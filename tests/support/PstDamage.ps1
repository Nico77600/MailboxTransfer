# Damaged-copy generators for the tests of the PST checker and of the repair (the same ones as tests\PstChecker.Tests.ps1):
# a clean PST written by our writer, and copies with a flipped byte, a broken CRC, a B-tree cycle, a removed contents row, a wrong
# count, a dangling NID, a truncated file, random damage. Dot-source this file inside BeforeAll (it needs the engine loaded).

function New-Msg {
    param([string]$Subject, [int]$AttachBytes = 0, [switch]$Embedded, [switch]$Unread)
    $P = [MailboxTransfer.FxProperty]
    $m = [MailboxTransfer.FxMessage]::new()
    $m.Properties.Add($P::FromString(0x001A, 'IPM.Note'))
    $m.Properties.Add($P::FromString(0x0037, $Subject))
    $m.Properties.Add($P::FromString(0x1000, ('body ' * 400)))
    $m.Properties.Add($P::FromInt32(0x0E07, $(if ($Unread) { 0 } else { 1 })))
    $m.Properties.Add($P::FromInt32(0x8001, 5).Named([guid]'11111111-2222-3333-4444-555555555555', [uint32]0x8501))
    $r = [MailboxTransfer.FxRecipient]::new()
    $r.Properties.Add($P::FromInt32(0x0C15, 1)); $r.Properties.Add($P::FromString(0x3001, 'Bob'))
    $m.Recipients.Add($r)
    if ($AttachBytes -gt 0) {
        $a = [MailboxTransfer.FxAttachment]::new()
        $b = [byte[]]::new($AttachBytes); [Random]::new(1).NextBytes($b)
        $a.Properties.Add($P::FromInt32(0x3705, 1)); $a.Properties.Add($P::FromBinary(0x3701, $b)); $a.Properties.Add($P::FromString(0x3707, 'f.bin'))
        $m.Attachments.Add($a)
    }
    if ($Embedded) {
        $e = [MailboxTransfer.FxMessage]::new()
        $e.Properties.Add($P::FromString(0x0037, 'Inner'))
        $ea = [MailboxTransfer.FxAttachment]::new()
        $ea.Properties.Add($P::FromInt32(0x3705, 5)); $ea.Properties.Add($P::Create(0x3701, 0x000D, [byte[]]@()))
        $ea.Embedded = $e
        $m.Attachments.Add($ea)
    }
    return $m
}

function New-TestPst {
    param([string]$Dir, [string]$Name, [string[]]$Faults = @())
    $path = Join-Path $Dir ($Name + '.pst')
    $o = [MailboxTransfer.PstWriterOptions]::new(); $o.Faults = $Faults
    $w = [MailboxTransfer.PstWriter]::Create($path, 'Checked', $o)
    $a = $w.CreateFolder($null, 'A', 'IPF.Note'); $b = $w.CreateFolder($a, 'B', 'IPF.Note'); $c = $w.CreateFolder($null, 'C', 'IPF.Note')
    foreach ($i in 1..40) { [void]$w.AddMessage($a, (New-Msg "a$i" -AttachBytes $(if ($i % 10 -eq 0) { 30000 } else { 0 }) -Unread:($i % 4 -eq 0))) }
    foreach ($i in 1..15) { [void]$w.AddMessage($b, (New-Msg "b$i" -Embedded:($i % 5 -eq 0))) }
    [void]$w.AddMessage($c, (New-Msg 'big' -AttachBytes 3MB))
    $w.Close()
    return $path
}

function Copy-Pst { param([string]$Source, [string]$Dir, [string]$Name) $d = Join-Path $Dir ($Name + '.pst'); Copy-Item -LiteralPath $Source -Destination $d; $d }
function Set-PageCrc { param([byte[]]$Raw, [long]$Ib) [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($Raw, [int]$Ib, 496)).CopyTo($Raw, $Ib + 500) }
function Set-HeaderCrc { param([byte[]]$Raw) [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($Raw, 8, 471)).CopyTo($Raw, 4); [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($Raw, 8, 516)).CopyTo($Raw, 524) }
function Get-LeafPage {
    param([byte[]]$Raw, [long]$RootIb, [int]$Which = 0)
    $ib = $RootIb
    while ($Raw[$ib + 491] -gt 0) { $ib = [BitConverter]::ToInt64($Raw, $ib + 16 + 24 * $Which) }
    $ib
}

# Offset of the data block of a top-level node
function Get-NodeBlock {
    param([string]$Path, [uint32]$Nid)
    $f = [MailboxTransfer.PstFile]::Open($Path)
    try { $n = $f.OpenNode($Nid); $e = [MailboxTransfer.PstBbtEntry]::new(); [void]$f.TryFindBlock($n.BidData, [ref]$e); [pscustomobject]@{ Ib = [long]$e.Ib; Bid = $e.Bid; Cb = [int]$e.Cb } } finally { $f.Dispose() }
}

# NIDs of the messages of a folder by name ('A', 'B', 'C')
function Get-FolderMessageNids {
    param([string]$Path, [string]$Folder)
    $r = [MailboxTransfer.PstReader]::Open($Path)
    try { $f = $r.AllFolders() | Where-Object Name -EQ $Folder; , @($r.MessageNids($f, $false)) } finally { $r.Dispose() }
}

# The data tree (XBLOCK) of the attachment data of the message in folder C (3 MB)
function Get-BigAttachmentTree {
    param([string]$Path)
    $f = [MailboxTransfer.PstFile]::Open($Path)
    $nid = (Get-FolderMessageNids $Path 'C')[0]
    $n = $f.OpenNode([uint32]$nid); $att = $null
    foreach ($sn in $n.SubnodeNids) { if (([MailboxTransfer.PstFormat]::NidType($sn)) -eq 5) { $att = $n.OpenSubnode($sn) } }
    $pc = [MailboxTransfer.PstPcReader]::new($att); $big = $null
    foreach ($e in $pc.Entries) { if ($e.Id -eq 0x3701) { $big = $att.OpenSubnode($e.Hnid) } }
    [pscustomobject]@{ File = $f; Node = $big }
}

# The leaf entry of a node in the NBT of raw file bytes, and its parent NID changed (page CRC recomputed)
function Find-NbtLeaf {
    param([byte[]]$Raw, [uint32]$Nid)
    $ib = [BitConverter]::ToInt64($Raw, 224)
    while ($true) {
        $c = $Raw[$ib + 488]
        if ($Raw[$ib + 491] -eq 0) {
            for ($i = 0; $i -lt $c; $i++) { if ([BitConverter]::ToUInt64($Raw, $ib + $i * 32) -eq $Nid) { return [pscustomobject]@{ Page = $ib; Off = $ib + $i * 32 } } }
            return $null
        }
        $sel = 0
        for ($i = 0; $i -lt $c; $i++) { if ([BitConverter]::ToUInt64($Raw, $ib + $i * 24) -le $Nid) { $sel = $i } }
        $ib = [BitConverter]::ToInt64($Raw, $ib + $sel * 24 + 16)
    }
}
function Set-NbtParent {
    param([byte[]]$Raw, [uint32]$Nid, [uint32]$Parent)
    $e = Find-NbtLeaf $Raw $Nid
    if (-not $e) { throw "NID $Nid not in the NBT" }
    [BitConverter]::GetBytes($Parent).CopyTo($Raw, $e.Off + 24)
    Set-PageCrc $Raw $e.Page
}
function Get-FolderNid {
    param([string]$Path, [string]$Name)
    $r = [MailboxTransfer.PstReader]::Open($Path)
    try { [uint32](($r.AllFolders() | Where-Object Name -EQ $Name).Nid) } finally { $r.Dispose() }
}

# The node entry of $Nid loses its pointers (data BID dangling, no subnode tree): the node exists in the NBT but its blocks are no longer
# reachable from it. The blocks stay in the file with nothing pointing at them, as when the leaf page of the NBT that held the entry is lost.
function Set-NbtPointers {
    param([byte[]]$Raw, [uint32]$Nid)
    $e = Find-NbtLeaf $Raw $Nid
    if (-not $e) { throw "NID $Nid not in the NBT" }
    [BitConverter]::GetBytes([uint64]0x7777770).CopyTo($Raw, $e.Off + 8)
    [BitConverter]::GetBytes([uint64]0).CopyTo($Raw, $e.Off + 16)
    Set-PageCrc $Raw $e.Page
}

# File offset of the data block of the first subnode of a node (the long values of a property context)
function Get-FirstSubnodeBlock {
    param([string]$Path, [uint32]$Nid)
    $f = [MailboxTransfer.PstFile]::Open($Path)
    try { $n = $f.OpenNode($Nid); $s = $f.ReadSubnodes($n.BidSub)[0]; $e = [MailboxTransfer.PstBbtEntry]::new(); [void]$f.TryFindBlock($s.BidData, [ref]$e); [long]$e.Ib } finally { $f.Dispose() }
}

# A damaged copy of the clean file $Good. Kinds: namemap-ptr, namemap-ptr-decoy (the entries of the name map and of the decoy map), namemap-block (header of its heap), namemap-data (a byte in it), namemap-byte (the byte at -Seed), namemap-two (two bytes), namemap-strings, flip-folder, flip-message, flip-attach, hdr-crc, page-crc-nbt, page-sig-bbt, nbt-order,
# folder-orphan, folder-moved, folder-loop, folder-self, msg-orphan, msg-orphan0, bid-dangling, highwater, unique, bidnextp, pmap, amap-unmarked, amap-bitorder, overlap, xblock-total, truncated, truncated-tail, zeros,
# random, tiny, empty, random-damage
function New-DamagedPst {
    param([string]$Good, [string]$Dir, [string]$Kind, [int]$Seed = 1, [string]$NidSource = $Good)
    $p = Copy-Pst $Good $Dir "dmg-$Kind-$Seed"
    $raw = [IO.File]::ReadAllBytes($p)
    switch ($Kind) {
        'flip-folder' { $b = Get-NodeBlock $p 0x122; $raw[$b.Ib + 10] = $raw[$b.Ib + 10] -bxor 0xFF }
        'flip-message' { $nid = (Get-FolderMessageNids $p 'A')[3]; $b = Get-NodeBlock $p ([uint32]$nid); $raw[$b.Ib + 20] = $raw[$b.Ib + 20] -bxor 0xFF }
        'flip-attach' {
            $t = Get-BigAttachmentTree $p
            try {
                $x = $t.File.ReadBlock($t.Node.BidData)       # the XBLOCK: its 100th entry
                $bid = [BitConverter]::ToUInt64($x, 8 + 8 * 100)
                $e2 = [MailboxTransfer.PstBbtEntry]::new(); [void]$t.File.TryFindBlock($bid, [ref]$e2)
                $raw[[long]$e2.Ib + 77] = $raw[[long]$e2.Ib + 77] -bxor 0x5A
            } finally { $t.File.Dispose() }
        }
        'namemap-ptr' { Set-NbtPointers $raw 0x61 }
        'namemap-ptr-decoy' { Set-NbtPointers $raw 0x61; Set-NbtPointers $raw ([MailboxTransfer.PstFormat]::Nid(1, 0x777)) }
        'namemap-block' { $b = Get-NodeBlock $p 0x61; $raw[$b.Ib] = $raw[$b.Ib] -bxor 0xFF; $raw[$b.Ib + 1] = $raw[$b.Ib + 1] -bxor 0xFF }       # the header of the heap: nothing can be read
        'namemap-data' { $b = Get-NodeBlock $p 0x61; $raw[$b.Ib + [int]($b.Cb * 0.6)] = $raw[$b.Ib + [int]($b.Cb * 0.6)] -bxor 0x55 }       # a byte in the middle of the heap: one entry or one hash is wrong
        'namemap-byte' { $b = Get-NodeBlock $p 0x61; $raw[$b.Ib + $Seed] = $raw[$b.Ib + $Seed] -bxor 0x55 }       # the byte at offset Seed of the heap of the map
        'namemap-two' { $b = Get-NodeBlock $p 0x61; foreach ($o in [int]($b.Cb * 0.3), [int]($b.Cb * 0.6)) { $raw[$b.Ib + $o] = $raw[$b.Ib + $o] -bxor 0x55 } }       # two wrong bytes: the CRC cannot tell which
        'namemap-strings' { $ib = Get-FirstSubnodeBlock $p 0x61; $raw[$ib + 40] = $raw[$ib + 40] -bxor 0xFF }
        'folder-orphan' { Set-NbtParent $raw (Get-FolderNid $NidSource 'B') 0x99999 }                       # parent node does not exist
        'folder-moved' { Set-NbtParent $raw (Get-FolderNid $NidSource 'B') (Get-FolderNid $NidSource 'C') }         # parent is another folder
        'folder-self' { Set-NbtParent $raw (Get-FolderNid $NidSource 'B') (Get-FolderNid $NidSource 'B') }                 # B is its own parent
        'folder-loop' { Set-NbtParent $raw (Get-FolderNid $NidSource 'A') (Get-FolderNid $NidSource 'B') }          # A under B under A
        'msg-orphan' { foreach ($n in (Get-FolderMessageNids $NidSource 'C')) { Set-NbtParent $raw ([uint32]$n) 0x8888 } }   # the message of C: folder does not exist
        'msg-orphan0' { foreach ($n in (Get-FolderMessageNids $NidSource 'C')) { Set-NbtParent $raw ([uint32]$n) 0 } }
        'hdr-crc' { $raw[20] = $raw[20] -bxor 1 }
        'page-crc-nbt' { $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 224)); $raw[$leaf + 5] = $raw[$leaf + 5] -bxor 0x55 }
        'page-sig-bbt' { $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 240)); $raw[$leaf + 498] = $raw[$leaf + 498] -bxor 1; Set-PageCrc $raw $leaf }
        'nbt-order' {
            $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 224))
            $a = $raw[$leaf..($leaf + 7)]; $b = $raw[($leaf + 32)..($leaf + 39)]
            $b.CopyTo($raw, $leaf); $a.CopyTo($raw, $leaf + 32); Set-PageCrc $raw $leaf
        }
        'bid-dangling' { $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 224)); [BitConverter]::GetBytes([uint64]0x7777770).CopyTo($raw, $leaf + 8); Set-PageCrc $raw $leaf }
        'highwater' { [BitConverter]::GetBytes([uint32]0x400).CopyTo($raw, 44 + 4 * 4); Set-HeaderCrc $raw }
        'unique' { [BitConverter]::GetBytes([uint32]0).CopyTo($raw, 40); Set-HeaderCrc $raw }
        'bidnextp' { [BitConverter]::GetBytes([uint64]1).CopyTo($raw, 32); Set-HeaderCrc $raw }
        'pmap' { $raw[0x4600 + 3] = 0x7F; Set-PageCrc $raw 0x4600 }
        'amap-bitorder' {
            for ($i = 0; $i -lt 496; $i++) { $b = $raw[0x4400 + $i]; $r = 0; for ($k = 0; $k -lt 8; $k++) { if ($b -band (1 -shl $k)) { $r = $r -bor (0x80 -shr $k) } }; $raw[0x4400 + $i] = $r }
            Set-PageCrc $raw 0x4400
        }
        'amap-unmarked' {
            $b = Get-NodeBlock $p 0x122; $ib = $b.Ib
            $base = 0x4400 + [math]::Floor(($ib - 0x4400) / 253952) * 253952
            $unit = [int](($ib - $base) / 64)
            $raw[$base + [math]::Floor($unit / 8)] = $raw[$base + [math]::Floor($unit / 8)] -band ((-bnot (0x80 -shr ($unit % 8))) -band 0xFF)
            Set-PageCrc $raw $base
        }
        'overlap' {
            $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 240))
            $raw[($leaf + 8)..($leaf + 15)].CopyTo($raw, $leaf + 24 + 8); Set-PageCrc $raw $leaf
        }
        'xblock-total' {
            $t = Get-BigAttachmentTree $p
            try {
                $e2 = [MailboxTransfer.PstBbtEntry]::new(); [void]$t.File.TryFindBlock($t.Node.BidData, [ref]$e2)
                $ib = [long]$e2.Ib
                $raw[$ib + 4] = $raw[$ib + 4] + 1
                $total = [MailboxTransfer.PstFormat]::BlockSize($e2.Cb)
                [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($raw, [int]$ib, $e2.Cb)).CopyTo($raw, $ib + $total - 16 + 4)
            } finally { $t.File.Dispose() }
        }
        'truncated' { $raw = $raw[0..([int]($raw.Length / 2))] }
        'truncated-tail' { $raw = $raw[0..($raw.Length - 40000)] }
        'zeros' { $raw = [byte[]]::new(300000) }
        'random' { $raw = [byte[]]::new(300000); [Random]::new($Seed).NextBytes($raw) }
        'tiny' { $raw = [byte[]](1..100) }
        'empty' { $raw = [byte[]]@() }
        'random-damage' {
            $rnd = [Random]::new($Seed)
            for ($k = 0; $k -lt 30; $k++) { $raw[$rnd.Next(0, $raw.Length)] = [byte]$rnd.Next(0, 256) }
        }
        default { throw "unknown damage kind $Kind" }
    }
    [IO.File]::WriteAllBytes($p, $raw)
    $p
}
