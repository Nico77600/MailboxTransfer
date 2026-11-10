#Requires -Version 7.4
<#
.SYNOPSIS
    Builds package\docs\MailboxTransfer-UserGuide.html and MailboxTransfer-Guide.html from the Markdown files next to them
    (or the one file given with -Source).

.DESCRIPTION
    The Markdown guide stays readable as plain text (and on GitHub / Azure DevOps). This script
    turns it into a structured, self-contained HTML page:

      - hero header (title, version, author, date) built from the front matter,
      - sticky sidebar with the parts and chapters, highlighting the chapter being read,
      - each chapter ("## ...") in its own card, with an icon and a number badge,
      - callouts from GitHub alerts (> [!NOTE], [!TIP], [!IMPORTANT], [!WARNING], [!CAUTION]),
      - four custom blocks written as fenced code in the Markdown:
          ```cards   icon | title | text          (grid of small cards)
          ```steps   title | text                 (numbered timeline)
          ```flow    icon | title | subtitle      (diagram; "arrow | label | text" = connector)
          ```gallery  images/file.png | caption   (2x2 responsive image gallery)
      - code blocks with a language label and a Copy button,
      - images embedded (the HTML file can be sent alone) with a zoom on click,
      - light / dark theme, print layout.

    No external resource is used (fonts, scripts, images are inline), so the page also works
    when opened from a mail attachment or a OneDrive / SharePoint preview.

    Section icons: put <!-- icon: name --> on the line before a "## " heading. Available names
    are the keys of $Icons below; add an SVG path there to add an icon.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
    PowerShell pitfall: never name a variable $matches — every -match overwrites the automatic
    $Matches, and variable names are case-insensitive.
#>
[CmdletBinding()]
param([string]$Source, [string]$Destination)
$ErrorActionPreference = 'Stop'
if (-not $Source) {
    foreach ($n in 'UserGuide', 'Guide') {
        & $PSCommandPath -Source (Join-Path $PSScriptRoot "..\package\docs\MailboxTransfer-$n.md")
    }
    return
}
if (-not $Destination) { $Destination = [IO.Path]::ChangeExtension((Resolve-Path $Source).Path, '.html') }
$Source = (Resolve-Path $Source).Path
$Destination = [IO.Path]::GetFullPath($Destination)
$docs = Split-Path $Source -Parent
$enc = { param($t) [System.Net.WebUtility]::HtmlEncode($t) }

# ---- Icons (24x24, stroke) ------------------------------------------------------------------
$Icons = @{
    book      = '<path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z"/><path d="M4 19V5"/><path d="M8 7h7M8 11h5"/>'
    flow      = '<rect x="3" y="4" width="7" height="5" rx="1.5"/><rect x="14" y="15" width="7" height="5" rx="1.5"/><path d="M6.5 9v4a2 2 0 0 0 2 2H14"/>'
    lightbulb = '<path d="M9 18h6M10 21h4"/><path d="M12 3a6 6 0 0 0-4 10.5c.7.7 1 1.5 1 2.5h6c0-1 .3-1.8 1-2.5A6 6 0 0 0 12 3z"/>'
    checklist = '<path d="M10 6h10M10 12h10M10 18h10"/><path d="m3.5 6 1.5 1.5L8 4.5M3.5 12l1.5 1.5L8 10.5M3.5 18l1.5 1.5L8 16.5"/>'
    download  = '<path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M4 20h16"/>'
    settings  = '<path d="M4 6h9M17 6h3M4 12h3M11 12h9M4 18h11M19 18h1"/><circle cx="15" cy="6" r="2"/><circle cx="9" cy="12" r="2"/><circle cx="17" cy="18" r="2"/>'
    clock     = '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>'
    terminal  = '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="m7 9 3 3-3 3M13 15h4"/>'
    chart     = '<path d="M3 20h18"/><path d="M6 16v-5M11 16V6M16 16v-8"/>'
    layers    = '<path d="m12 3 9 5-9 5-9-5z"/><path d="m3 13 9 5 9-5"/>'
    gear      = '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.9 4.9 7 7M17 17l2.1 2.1M2 12h3M19 12h3M4.9 19.1 7 17M17 7l2.1-2.1"/>'
    wrench    = '<path d="M15 4a5 5 0 0 0-4.6 6.9L3 18.3 5.7 21l7.4-7.4A5 5 0 0 0 20 9l-3 1-3-3 1-3z"/>'
    beaker    = '<path d="M9 3h6M10 3v6L4.5 18.5A1.7 1.7 0 0 0 6 21h12a1.7 1.7 0 0 0 1.5-2.5L14 9V3"/><path d="M7 15h10"/>'
    lifebuoy  = '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="4"/><path d="m5.6 5.6 3.6 3.6M14.8 14.8l3.6 3.6M18.4 5.6l-3.6 3.6M9.2 14.8l-3.6 3.6"/>'
    compare   = '<rect x="3" y="4" width="7" height="16" rx="1.5"/><rect x="14" y="4" width="7" height="16" rx="1.5"/><path d="M6.5 9h0M17.5 9h0"/>'
    database  = '<ellipse cx="12" cy="5.5" rx="8" ry="3"/><path d="M4 5.5v13c0 1.7 3.6 3 8 3s8-1.3 8-3v-13"/><path d="M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>'
    tag       = '<path d="M3 12V4h8l10 10-8 8z"/><circle cx="7.5" cy="8" r="1.5"/>'
    target    = '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1"/>'
    file      = '<path d="M14 3H6v18h12V7z"/><path d="M14 3v4h4M9 13h6M9 17h6"/>'
    calendar  = '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>'
    refresh   = '<path d="M20 11a8 8 0 0 0-14.3-4.9L4 8"/><path d="M4 4v4h4"/><path d="M4 13a8 8 0 0 0 14.3 4.9L20 16"/><path d="M20 20v-4h-4"/>'
    people    = '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0"/><circle cx="17" cy="9" r="2.5"/><path d="M16 14.5a5 5 0 0 1 5.5 5"/>'
    key       = '<circle cx="8" cy="15" r="4"/><path d="m11 12 9-9M16 7l3 3"/>'
    shield    = '<path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="m9 12 2 2 4-4"/>'
    info      = '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.5"/>'
    search    = '<circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/>'
    check     = '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>'
    play      = '<circle cx="12" cy="12" r="9"/><path d="m10 8 6 4-6 4z"/>'
    mail      = '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/>'
    copy      = '<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1v10a1 1 0 0 0 1 1h3"/>'
    moon      = '<path d="M20 14.5A8 8 0 0 1 9.5 4 8 8 0 1 0 20 14.5z"/>'
    up        = '<path d="m6 14 6-6 6 6"/>'
    user      = '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>'
}
function Get-Icon([string]$Name, [string]$Class = 'icon') {
    $path = $Icons[$Name]; if (-not $path) { $path = $Icons['info'] }
    "<svg class=""$Class"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">$path</svg>"
}
function ConvertTo-Inline([string]$Text) {
    # Inline Markdown (code, bold, links) inside a custom block field.
    $h = (ConvertFrom-Markdown -InputObject $Text.Trim()).Html.Trim()
    return ($h -replace '^<p>', '' -replace '</p>$', '')
}
function Resolve-DocImageSource([string]$RelativePath) {
    $decoded = [Uri]::UnescapeDataString($RelativePath)
    $path = if ([IO.Path]::IsPathRooted($decoded)) { $decoded } else { Join-Path $docs $decoded }
    if (-not (Test-Path -LiteralPath $path)) { return $RelativePath }
    $mime = switch ([IO.Path]::GetExtension($path).ToLowerInvariant()) {
        '.jpg' { 'image/jpeg'; break }
        '.jpeg' { 'image/jpeg'; break }
        '.gif' { 'image/gif'; break }
        '.webp' { 'image/webp'; break }
        '.svg' { 'image/svg+xml'; break }
        default { 'image/png' }
    }
    return "data:$mime;base64,$([Convert]::ToBase64String([IO.File]::ReadAllBytes($path)))"
}

# ---- Front matter --------------------------------------------------------------------------------
$markdown = [IO.File]::ReadAllText($Source) -replace "`r`n", "`n"
$meta = @{ title = 'Guide'; subtitle = ''; version = ''; author = ''; updated = '' }
$front = [regex]::Match($markdown, '(?s)\A---\n(.*?)\n---\n')
if ($front.Success) {
    foreach ($line in $front.Groups[1].Value -split "`n") { if ($line -match '^\s*(\w+)\s*:\s*(.*)$') { $meta[$Matches[1]] = $Matches[2].Trim() } }
    $markdown = $markdown.Substring($front.Length)
}
$html = (ConvertFrom-Markdown -InputObject $markdown).Html

# ---- Custom blocks -----------------------------------------------------------------------------------
$html = [regex]::Replace($html, '(?s)<pre><code class="language-(cards|steps|flow|gallery)">(.*?)</code></pre>', {
        param($m)
        $kind = $m.Groups[1].Value
        $lines = [System.Net.WebUtility]::HtmlDecode($m.Groups[2].Value) -split "`n" | Where-Object { $_.Trim() }
        switch ($kind) {
            'cards' {
                $items = foreach ($l in $lines) {
                    $icon, $title, $text = $l.Split('|', 3).ForEach({ $_.Trim() })
                    "<div class=""card-item""><div class=""card-icon"">$(Get-Icon $icon)</div><div><div class=""card-title"">$(ConvertTo-Inline $title)</div><div class=""card-text"">$(ConvertTo-Inline $text)</div></div></div>"
                }
                "<div class=""cards"">$($items -join '')</div>"
            }
            'steps' {
                $n = 0
                $items = foreach ($l in $lines) {
                    $n++
                    $title, $text = $l.Split('|', 2).ForEach({ $_.Trim() })
                    "<li><span class=""step-num"">$n</span><div class=""step-body""><div class=""step-title"">$(ConvertTo-Inline $title)</div><div class=""step-text"">$(ConvertTo-Inline $text)</div></div></li>"
                }
                "<ol class=""steps"">$($items -join '')</ol>"
            }
            'flow' {
                $items = foreach ($l in $lines) {
                    $icon, $title, $sub = $l.Split('|', 3).ForEach({ $_.Trim() })
                    if ($icon -eq 'arrow') {
                        "<div class=""flow-arrow""><span class=""flow-label"">$(& $enc $title)</span><svg viewBox=""0 0 40 12"" aria-hidden=""true""><path d=""M0 6h36M31 1l6 5-6 5"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-sub"">$(& $enc $sub)</span></div>"
                    } else {
                        "<div class=""flow-node""><div class=""flow-icon"">$(Get-Icon $icon)</div><div class=""flow-title"">$(& $enc $title)</div><div class=""flow-text"">$(& $enc $sub)</div></div>"
                    }
                }
                "<div class=""flow"">$($items -join '')</div>"
            }
            'gallery' {
                $items = foreach ($l in $lines) {
                    $imagePath, $caption = $l.Split('|', 2).ForEach({ $_.Trim() })
                    if (-not $caption) { $caption = [IO.Path]::GetFileNameWithoutExtension($imagePath) }
                    $src = Resolve-DocImageSource $imagePath
                    "<figure class=""gallery-item""><img loading=""lazy"" src=""$src"" alt=""$(& $enc $caption)"" title=""Click to enlarge""><figcaption>$(ConvertTo-Inline $caption)</figcaption></figure>"
                }
                "<div class=""gallery"">$($items -join '')</div>"
            }
        }
    })

# ---- Code blocks: label + copy button ------------------------------------------------------------------
$languages = @{ powershell = 'PowerShell'; sql = 'SQL'; text = 'Text'; json = 'JSON'; '' = 'Code' }
$html = [regex]::Replace($html, '(?s)<pre><code(?: class="language-([\w-]+)")?>(.*?)</code></pre>', {
        param($m)
        $lang = $m.Groups[1].Value
        $label = if ($languages.ContainsKey($lang)) { $languages[$lang] } else { $lang }
        "<div class=""code""><div class=""code-head""><span>$label</span><button type=""button"" class=""copy"">$(Get-Icon 'copy' 'icon-sm')<span>Copy</span></button></div><pre><code>$($m.Groups[2].Value)</code></pre></div>"
    })

# ---- Tables, images, links, alerts -------------------------------------------------------------------------
$html = $html.Replace('<table>', '<div class="table-wrap"><table>').Replace('</table>', '</table></div>')
$html = [regex]::Replace($html, '<p><img src="([^"]+)" alt="([^"]*)"\s*/?></p>', {
        param($m)
        $caption = [System.Net.WebUtility]::HtmlDecode($m.Groups[2].Value)
        $src = Resolve-DocImageSource $m.Groups[1].Value
        "<figure><img loading=""lazy"" src=""$src"" alt=""$(& $enc $caption)"" title=""Click to enlarge""><figcaption>$(& $enc $caption)</figcaption></figure>"
    })
$html = $html -replace '<a href="(https?://[^"]+)"', '<a href="$1" target="_blank" rel="noopener"'
$html = $html -replace '<p class="markdown-alert-title"><svg viewBox="0 0 16 16"', '<p class="markdown-alert-title"><svg class="icon-sm" viewBox="0 0 16 16" fill="currentColor"'

# ---- Parts and chapters ------------------------------------------------------------------------------------
$pattern = '(?s)(?:<!-- icon: (?<icon>[\w-]+) -->\s*)?<h(?<lvl>[12]) id="(?<id>[^"]*)">(?<text>.*?)</h\k<lvl>>'
$headings = [regex]::Matches($html, $pattern)
$hero = ''; $body = [Text.StringBuilder]::new(); $nav = [Text.StringBuilder]::new()
$sectionOpen = $false
for ($i = 0; $i -lt $headings.Count; $i++) {
    $m = $headings[$i]
    $end = if ($i + 1 -lt $headings.Count) { $headings[$i + 1].Index } else { $html.Length }
    $content = $html.Substring($m.Index + $m.Length, $end - $m.Index - $m.Length)
    $text = $m.Groups['text'].Value; $id = $m.Groups['id'].Value
    if ($i -eq 0 -and $m.Groups['lvl'].Value -eq '1') {
        $lead = ''
        $quote = [regex]::Match($content, '(?s)^\s*<blockquote>\s*<p>(.*?)</p>\s*</blockquote>')
        if ($quote.Success) { $lead = $quote.Groups[1].Value; $content = $content.Substring($quote.Length) }
        $badges = @(
            "<span class=""badge badge-accent"">Version $($meta.version)</span>",
            "<span class=""badge"">$(Get-Icon 'calendar' 'icon-sm')Updated $($meta.updated)</span>",
            "<span class=""badge"">$(Get-Icon 'user' 'icon-sm')$($meta.author)</span>",
            "<span class=""badge"">$(Get-Icon 'terminal' 'icon-sm')Collector 5.1 · Analysis 7.4+</span>",
            "<span class=""badge"">$(Get-Icon 'shield' 'icon-sm')Read-only collector</span>") -join ''
        $hero = "<header class=""hero"" id=""top""><div class=""hero-top""><div class=""hero-logo"">$(Get-Icon 'shield')</div><div><div class=""eyebrow"">$(& $enc $meta.subtitle)</div><h1>$(& $enc $meta.title)</h1></div>" +
            "<button type=""button"" id=""theme"" class=""ghost"" title=""Light / dark"">$(Get-Icon 'moon' 'icon-sm')</button></div>" +
            "<div class=""badges"">$badges</div><p class=""lead"">$lead</p>$content</header>"
        continue
    }
    if ($sectionOpen) { [void]$body.Append('</div></section>'); $sectionOpen = $false }
    if ($m.Groups['lvl'].Value -eq '1') {
        $label, $name = ($text -split ' · ', 2)
        $partName = if ($name) { '<span class="part-name">' + $name + '</span>' } else { '' }
        $navName = if ($name) { ' <span>' + $name + '</span>' } else { '' }
        [void]$body.Append("<div class=""part"" id=""$id""><span class=""part-label"">$label</span>$partName</div>")
        [void]$nav.Append("<div class=""nav-part"">$label$navName</div>")
        [void]$body.Append($content)
        continue
    }
    $icon = if ($m.Groups['icon'].Success) { $m.Groups['icon'].Value } else { 'play' }
    $badge = ''; $eyebrow = ''; $title = $text
    $numbered = [regex]::Match($text, '^(\d+)\.\s+(.*)$')
    $annex = [regex]::Match($text, '^Annex ([A-Z]) — (.*)$')
    if ($numbered.Success) { $badge = $numbered.Groups[1].Value; $title = $numbered.Groups[2].Value }
    elseif ($annex.Success) { $badge = $annex.Groups[1].Value; $title = $annex.Groups[2].Value; $eyebrow = "Annex $badge" }
    $eyebrowHtml = if ($eyebrow) { "<div class=""section-eyebrow"">$eyebrow</div>" } elseif ($badge) { "<div class=""section-eyebrow"">Chapter $badge</div>" } else { "<div class=""section-eyebrow"">Start here</div>" }
    [void]$body.Append("<section class=""section"" id=""$id""><div class=""section-head""><div class=""section-icon"">$(Get-Icon $icon)</div><div>$eyebrowHtml<h2>$title</h2></div></div><div class=""section-body"">")
    [void]$body.Append($content)
    $sectionOpen = $true
    $navBadge = if ($badge) { $badge } else { Get-Icon 'play' 'icon-xs' }
    [void]$nav.Append("<a href=""#$id"" data-target=""$id""><span class=""nav-badge"">$navBadge</span><span class=""nav-text"">$title</span></a>")
}
if ($sectionOpen) { [void]$body.Append('</div></section>') }

# ---- Page ------------------------------------------------------------------------------------------------------
$page = @"
<!doctype html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(& $enc $meta.title) - $(& $enc $meta.subtitle)</title>
<script>
  (() => {
    const param = new URLSearchParams(window.location.search).get("scoutTheme");
    const theme = param || (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
    document.documentElement.setAttribute("data-theme", theme);
  })();
</script>
<style>
:root {
  color-scheme: light;
  --cp-bg: #f7f4ef; --cp-bg-elevated: #fcfbf8; --cp-surface: #ffffff; --cp-surface-soft: #f5f5f5;
  --cp-border: #dedede; --cp-border-strong: #919191; --cp-text: #242424; --cp-text-muted: #5c5c5c; --cp-text-soft: #6f6f6f;
  --cp-accent: #b11f4b; --cp-accent-hover: #9a1a41; --cp-accent-soft: rgba(177, 31, 75, 0.08); --cp-accent-fg: #ffffff;
  --cp-success: #16a34a; --cp-danger: #dc2626; --cp-warning: #f59e0b; --cp-link: #0078d4;
  --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.12); --cp-overlay: rgba(255, 255, 255, 0.8);
  --cp-panel: rgba(255, 255, 255, 0.86); --cp-panel-strong: rgba(255, 255, 255, 0.96); --cp-sheen: rgba(255, 255, 255, 0.55);
  --cp-highlight: rgba(177, 31, 75, 0.12);
}
html[data-theme="dark"] {
  color-scheme: dark;
  --cp-bg: #3d3b3a; --cp-bg-elevated: #343231; --cp-surface: #292929; --cp-surface-soft: #2e2e2e;
  --cp-border: #474747; --cp-border-strong: #5f5f5f; --cp-text: #dedede; --cp-text-muted: #919191; --cp-text-soft: #b0b0b0;
  --cp-accent: #fd8ea1; --cp-accent-hover: #fb7b91; --cp-accent-soft: rgba(253, 142, 161, 0.14); --cp-accent-fg: #1a1a1a;
  --cp-success: #4ade80; --cp-danger: #f87171; --cp-warning: #fbbf24; --cp-link: #4da6ff;
  --cp-shadow: 0 18px 48px rgba(0, 0, 0, 0.32); --cp-overlay: rgba(41, 41, 41, 0.88);
  --cp-panel: rgba(41, 41, 41, 0.72); --cp-panel-strong: rgba(41, 41, 41, 0.96); --cp-sheen: rgba(255, 255, 255, 0.04);
  --cp-highlight: rgba(253, 142, 161, 0.12);
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; scroll-padding-top: 24px; }
body { margin: 0; background: var(--cp-bg); color: var(--cp-text); font: 15.5px/1.7 "Segoe UI", Aptos, Calibri, -apple-system, BlinkMacSystemFont, sans-serif; display: grid; grid-template-columns: 300px minmax(0, 1fr); }
a { color: var(--cp-link); text-decoration: none; } a:hover { text-decoration: underline; }
.icon { width: 22px; height: 22px; } .icon-sm { width: 16px; height: 16px; flex-shrink: 0; } .icon-xs { width: 12px; height: 12px; }
button { font: inherit; color: var(--cp-text); background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: 0.625rem; cursor: pointer; }
button:hover { border-color: var(--cp-accent); color: var(--cp-accent); }
:focus-visible { outline: 2px solid var(--cp-accent); outline-offset: 2px; }

/* ---- Sidebar -------------------------------------------------------------------- */
aside { position: sticky; top: 0; height: 100vh; overflow-y: auto; background: var(--cp-bg-elevated); border-right: 1px solid var(--cp-border); padding: 24px 16px 32px; }
.brand { display: flex; gap: 12px; align-items: center; padding: 4px 8px 20px; border-bottom: 1px solid var(--cp-border); margin-bottom: 12px; }
.brand-logo { width: 40px; height: 40px; border-radius: 12px; display: grid; place-items: center; background: var(--cp-accent); color: var(--cp-accent-fg); }
.brand-name { font-weight: 700; font-size: 15px; line-height: 1.2; } .brand-sub { font-size: 12px; color: var(--cp-text-muted); }
.nav-part { margin: 20px 8px 6px; font-size: 11px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--cp-accent); }
.nav-part span { color: var(--cp-text-muted); font-weight: 600; letter-spacing: 0.04em; }
aside a { display: flex; align-items: center; gap: 10px; padding: 6px 8px; margin: 1px 0; border-radius: 0.625rem; color: var(--cp-text); font-size: 13.5px; line-height: 1.35; border-left: 3px solid transparent; }
aside a:hover { background: var(--cp-accent-soft); text-decoration: none; }
aside a.active { background: var(--cp-accent-soft); border-left-color: var(--cp-accent); font-weight: 600; }
.nav-badge { flex-shrink: 0; width: 24px; height: 24px; border-radius: 50%; display: grid; place-items: center; font-size: 11.5px; font-weight: 700; background: var(--cp-surface-soft); color: var(--cp-text-muted); border: 1px solid var(--cp-border); }
aside a.active .nav-badge { background: var(--cp-accent); color: var(--cp-accent-fg); border-color: var(--cp-accent); }

/* ---- Main -------------------------------------------------------------------------- */
main { padding: 40px 56px 80px; max-width: 1080px; width: 100%; }
.hero { position: relative; overflow: hidden; border: 1px solid var(--cp-border); border-top: 4px solid var(--cp-accent); border-radius: 16px; padding: 36px 40px 32px; margin-bottom: 40px; background: linear-gradient(125deg, var(--cp-surface) 40%, var(--cp-accent-soft)); }
.hero::before { content: ""; position: absolute; width: 320px; height: 320px; right: -110px; top: -170px; border: 44px solid var(--cp-highlight); border-radius: 50%; pointer-events: none; }
.hero-top { position: relative; display: flex; align-items: center; gap: 18px; }
.hero-logo { width: 60px; height: 60px; border-radius: 16px; display: grid; place-items: center; background: var(--cp-accent); color: var(--cp-accent-fg); flex-shrink: 0; }
.hero-logo .icon { width: 32px; height: 32px; }
.eyebrow, .section-eyebrow { font-size: 11.5px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--cp-accent); }
.hero h1 { font-size: 38px; font-weight: 700; letter-spacing: -0.025em; line-height: 1.1; margin: 4px 0 0; }
#theme { margin-left: auto; padding: 8px 10px; display: grid; place-items: center; }
.badges { position: relative; display: flex; flex-wrap: wrap; gap: 8px; margin: 22px 0 18px; }
.badge { display: inline-flex; align-items: center; gap: 6px; padding: 4px 12px; border-radius: 999px; font-size: 12.5px; background: var(--cp-surface); border: 1px solid var(--cp-border); color: var(--cp-text-muted); }
.badge-accent { background: var(--cp-accent); color: var(--cp-accent-fg); border-color: var(--cp-accent); font-weight: 600; }
.lead { position: relative; font-size: 17.5px; line-height: 1.6; max-width: 760px; margin: 0 0 24px; }

.part { display: flex; align-items: baseline; gap: 14px; margin: 56px 0 20px; padding-bottom: 10px; border-bottom: 2px solid var(--cp-accent); }
.part-label { font-size: 13px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--cp-accent); }
.part-name { font-size: 28px; font-weight: 700; letter-spacing: -0.02em; }

.section { background: var(--cp-surface); border: 1px solid var(--cp-border); border-radius: 16px; padding: 32px 40px 28px; margin: 0 0 28px; box-shadow: 0 0 2px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.14); }
.section-head { display: flex; align-items: center; gap: 16px; padding-bottom: 18px; margin-bottom: 8px; border-bottom: 1px solid var(--cp-border); }
.section-icon { width: 48px; height: 48px; border-radius: 14px; display: grid; place-items: center; background: var(--cp-accent-soft); color: var(--cp-accent); flex-shrink: 0; }
.section-icon .icon { width: 26px; height: 26px; }
.section h2 { font-size: 25px; font-weight: 700; letter-spacing: -0.02em; margin: 2px 0 0; line-height: 1.2; }
.section-body > :first-child { margin-top: 16px; }
h3 { font-size: 17.5px; font-weight: 650; margin: 36px 0 12px; padding-left: 12px; border-left: 3px solid var(--cp-accent); line-height: 1.3; }
p { margin: 12px 0; } ul, ol { padding-left: 22px; } li { margin: 6px 0; } li::marker { color: var(--cp-accent); }
strong { font-weight: 650; }
code { font-family: Consolas, "Courier New", Courier, monospace; font-size: 0.88em; background: var(--cp-surface-soft); border: 1px solid var(--cp-border); padding: 1px 6px; border-radius: 6px; }
kbd { font-family: Consolas, monospace; font-size: 0.85em; padding: 1px 7px; border: 1px solid var(--cp-border-strong); border-bottom-width: 2px; border-radius: 6px; background: var(--cp-surface-soft); }
hr { border: 0; border-top: 1px solid var(--cp-border); margin: 28px 0; }

/* Tables */
.table-wrap { overflow-x: auto; margin: 16px 0 20px; border: 1px solid var(--cp-border); border-radius: 12px; }
table { width: 100%; border-collapse: collapse; font-size: 14px; line-height: 1.55; }
th { text-align: left; font-weight: 650; font-size: 12.5px; letter-spacing: 0.02em; color: var(--cp-text-muted); background: var(--cp-surface-soft); padding: 11px 16px; border-bottom: 1px solid var(--cp-border); }
td { padding: 11px 16px; border-bottom: 1px solid var(--cp-border); vertical-align: top; }
td:first-child { font-weight: 600; }
tr:last-child td { border-bottom: 0; }
tbody tr:hover td { background: var(--cp-accent-soft); }
td code { white-space: nowrap; }

/* Code */
.code { margin: 16px 0 20px; border: 1px solid var(--cp-border); border-radius: 12px; overflow: hidden; background: var(--cp-surface-soft); }
.code-head { display: flex; justify-content: space-between; align-items: center; padding: 6px 8px 6px 16px; border-bottom: 1px solid var(--cp-border); font-size: 12px; font-weight: 600; color: var(--cp-text-muted); letter-spacing: 0.04em; text-transform: uppercase; }
.copy { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; font-size: 12px; text-transform: none; letter-spacing: 0; }
.copy.done { color: var(--cp-success); border-color: var(--cp-success); }
.code pre { margin: 0; padding: 16px 18px; overflow-x: auto; font-size: 13.5px; line-height: 1.6; }
.code pre code { background: none; border: 0; padding: 0; font-size: inherit; }

/* Callouts */
.markdown-alert { --tone: var(--cp-link); margin: 18px 0; padding: 14px 18px; border: 1px solid color-mix(in srgb, var(--tone) 35%, var(--cp-border)); border-left: 4px solid var(--tone); border-radius: 12px; background: color-mix(in srgb, var(--tone) 7%, var(--cp-surface)); }
.markdown-alert p { margin: 6px 0; }
.markdown-alert-title { display: flex; align-items: center; gap: 8px; font-weight: 700; font-size: 13px; letter-spacing: 0.04em; text-transform: uppercase; color: var(--tone); margin: 0 0 4px !important; }
.markdown-alert-tip { --tone: var(--cp-success); } .markdown-alert-important { --tone: var(--cp-accent); }
.markdown-alert-warning { --tone: var(--cp-warning); } .markdown-alert-caution { --tone: var(--cp-danger); }

/* Cards */
.cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 14px; margin: 18px 0 22px; position: relative; }
.card-item { display: flex; gap: 14px; padding: 16px 18px; border: 1px solid var(--cp-border); border-radius: 14px; background: var(--cp-surface); transition: transform 160ms ease, border-color 160ms ease; }
.card-item:hover { transform: translateY(-2px); border-color: var(--cp-accent); }
.card-icon { width: 38px; height: 38px; border-radius: 11px; display: grid; place-items: center; background: var(--cp-accent-soft); color: var(--cp-accent); flex-shrink: 0; }
.card-title { font-weight: 700; font-size: 14.5px; line-height: 1.3; margin: 2px 0 4px; }
.card-text { font-size: 13.5px; line-height: 1.5; color: var(--cp-text-muted); }

/* Steps */
.steps { list-style: none; padding: 0; margin: 18px 0 22px; position: relative; }
.steps li { position: relative; display: flex; gap: 16px; margin: 0; padding: 0 0 18px; }
.steps li:not(:last-child)::before { content: ""; position: absolute; left: 16px; top: 36px; bottom: 2px; width: 2px; background: var(--cp-border); }
.step-num { flex-shrink: 0; width: 34px; height: 34px; border-radius: 50%; display: grid; place-items: center; font-weight: 700; font-size: 14px; background: var(--cp-accent); color: var(--cp-accent-fg); box-shadow: 0 0 0 4px var(--cp-accent-soft); }
.step-body { padding-top: 5px; } .step-title { font-weight: 700; font-size: 15px; line-height: 1.35; }
.step-text { color: var(--cp-text-muted); font-size: 14.5px; margin-top: 2px; }

/* Flow */
.flow { display: flex; align-items: stretch; flex-wrap: wrap; gap: 6px; margin: 20px 0 24px; padding: 20px; border-radius: 14px; background: var(--cp-surface-soft); border: 1px dashed var(--cp-border-strong); }
.flow-node { flex: 1 1 130px; min-width: 120px; text-align: center; padding: 16px 12px; border-radius: 12px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.flow-icon { width: 42px; height: 42px; margin: 0 auto 8px; border-radius: 12px; display: grid; place-items: center; background: var(--cp-accent-soft); color: var(--cp-accent); }
.flow-title { font-weight: 700; font-size: 14px; line-height: 1.3; } .flow-text { font-size: 12.5px; color: var(--cp-text-muted); line-height: 1.4; margin-top: 4px; }
.flow-arrow { flex: 0 0 auto; min-width: 70px; display: flex; flex-direction: column; align-items: center; justify-content: center; color: var(--cp-border-strong); padding: 0 2px; }
.flow-arrow svg { width: 44px; height: 14px; color: var(--cp-accent); }
.flow-label { font-size: 12px; font-weight: 700; color: var(--cp-accent); } .flow-sub { font-size: 11px; color: var(--cp-text-muted); text-align: center; max-width: 110px; line-height: 1.3; }

/* Figures */
figure { margin: 18px 0 24px; }
figure img { display: block; max-width: 100%; border: 1px solid var(--cp-border); border-radius: 12px; box-shadow: 0 0 2px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.14); cursor: zoom-in; }
figcaption { font-size: 13px; color: var(--cp-text-muted); margin-top: 8px; text-align: center; }
.gallery { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; margin: 20px 0 26px; }
.gallery figure { margin: 0; min-width: 0; }
.gallery img { width: 100%; background: var(--cp-surface-soft); }
.gallery figcaption { font-size: 12.5px; }
dialog#zoom { border: 0; padding: 0; background: transparent; max-width: 96vw; max-height: 96vh; }
dialog#zoom::backdrop { background: rgba(0, 0, 0, 0.7); }
dialog#zoom img { max-width: 96vw; max-height: 94vh; border-radius: 12px; cursor: zoom-out; display: block; }

#to-top { position: fixed; right: 24px; bottom: 24px; width: 44px; height: 44px; border-radius: 50%; display: grid; place-items: center; background: var(--cp-accent); color: var(--cp-accent-fg); border: 0; box-shadow: var(--cp-shadow); opacity: 0; pointer-events: none; transition: opacity 200ms; }
#to-top.visible { opacity: 1; pointer-events: auto; } #to-top:hover { background: var(--cp-accent-hover); color: var(--cp-accent-fg); }
footer { margin-top: 48px; padding-top: 20px; border-top: 1px solid var(--cp-border); font-size: 12.5px; color: var(--cp-text-muted); text-align: center; }

@media (max-width: 1000px) {
  body { display: block; }
  aside { position: static; height: auto; border-right: 0; border-bottom: 1px solid var(--cp-border); }
  main { padding: 24px 16px 60px; } .section { padding: 24px 20px; } .hero { padding: 28px 22px; }
  .gallery { grid-template-columns: 1fr; }
}
@media print {
  body { display: block; background: #fff; } aside, #to-top, #theme, .copy { display: none !important; }
  main { padding: 0; max-width: none; } .section { break-inside: avoid-page; box-shadow: none; } .part { break-before: page; }
}
</style></head><body>
<aside><div class="brand"><div class="brand-logo">$(Get-Icon 'shield')</div><div><div class="brand-name">$(& $enc $meta.title)</div><div class="brand-sub">$(& $enc $meta.subtitle) · v$($meta.version)</div></div></div>
<nav aria-label="Contents">$($nav.ToString())</nav></aside>
<main>$hero$($body.ToString())
<footer>$(& $enc $meta.title) $($meta.version) · $(& $enc $meta.author) · generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') from $([IO.Path]::GetFileName($Source))</footer></main>
<button type="button" id="to-top" title="Back to top">$(Get-Icon 'up' 'icon-sm')</button>
<dialog id="zoom"><img alt=""></dialog>
<script>
(() => {
  "use strict";
  // Theme toggle (no storage: the choice lasts for the page).
  document.getElementById("theme").addEventListener("click", () => {
    const root = document.documentElement;
    root.setAttribute("data-theme", root.getAttribute("data-theme") === "dark" ? "light" : "dark");
  });
  // Highlight the chapter being read in the sidebar.
  const links = new Map(Array.from(document.querySelectorAll("aside a[data-target]")).map(a => [a.dataset.target, a]));
  const observer = new IntersectionObserver(entries => {
    for (const e of entries) {
      if (!e.isIntersecting) continue;
      links.forEach(a => a.classList.remove("active"));
      const link = links.get(e.target.id);
      if (link) { link.classList.add("active"); }
    }
  }, { rootMargin: "-20% 0px -70% 0px" });
  document.querySelectorAll("section.section").forEach(s => observer.observe(s));
  // Copy buttons.
  document.querySelectorAll(".code .copy").forEach(button => button.addEventListener("click", async () => {
    const text = button.closest(".code").querySelector("pre").innerText;
    try { await navigator.clipboard.writeText(text); }
    catch { const r = document.createRange(); r.selectNodeContents(button.closest(".code").querySelector("pre")); const s = getSelection(); s.removeAllRanges(); s.addRange(r); document.execCommand("copy"); s.removeAllRanges(); }
    const label = button.querySelector("span"); label.textContent = "Copied"; button.classList.add("done");
    setTimeout(() => { label.textContent = "Copy"; button.classList.remove("done"); }, 1600);
  }));
  // Image zoom.
  const zoom = document.getElementById("zoom");
  document.querySelectorAll("figure img").forEach(img => img.addEventListener("click", () => { zoom.querySelector("img").src = img.src; zoom.showModal(); }));
  zoom.addEventListener("click", () => zoom.close());
  // Back to top.
  const top = document.getElementById("to-top");
  addEventListener("scroll", () => top.classList.toggle("visible", scrollY > 600), { passive: true });
  top.addEventListener("click", () => scrollTo({ top: 0 }));
})();
</script>
</body></html>
"@
[IO.File]::WriteAllText($Destination, $page, [Text.UTF8Encoding]::new($false))
Write-Host "Written: $Destination ($([Math]::Round((Get-Item $Destination).Length / 1KB)) KB)"
