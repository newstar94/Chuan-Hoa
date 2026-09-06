param([Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$file = [IO.File]::Open($OutputPath, 'CreateNew')
$zip = [IO.Compression.ZipArchive]::new($file, [IO.Compression.ZipArchiveMode]::Create)
function Write-ZipText([string]$Name, [string]$Text) {
    $writer = [IO.StreamWriter]::new($zip.CreateEntry($Name).Open(), [Text.UTF8Encoding]::new($false))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}
try {
    Write-ZipText '[Content_Types].xml' '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/></Types>'
    Write-ZipText '_rels/.rels' '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="r2" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI.xml"/></Relationships>'
    Write-ZipText 'word/document.xml' '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Ribbon alignment probe — no macros, no user data.</w:t></w:r></w:p></w:body></w:document>'
    Write-ZipText 'customUI/customUI.xml' '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="alignmentProbe" label="Alignment QA"><group id="test" label="Reset alignment"><box id="column" boxStyle="vertical"><button id="percent" label="100%" size="normal" showLabel="false" showImage="true" image="percent"/><button id="normal" label="Normal" size="normal" showLabel="false" showImage="true" image="dot"/></box></group></tab></tabs></ribbon></customUI>'
    Write-ZipText 'customUI/_rels/customUI.xml.rels' '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="percent" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="percent.png"/><Relationship Id="dot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="dot.png"/></Relationships>'
    foreach ($name in @('percent','dot')) {
        $bitmap = [Drawing.Bitmap]::new(36,16)
        $g = [Drawing.Graphics]::FromImage($bitmap)
        $font = [Drawing.Font]::new('Segoe UI',12,[Drawing.FontStyle]::Regular,[Drawing.GraphicsUnit]::Pixel)
        $format = [Drawing.StringFormat]::new()
        $format.Alignment = [Drawing.StringAlignment]::Center
        $format.LineAlignment = [Drawing.StringAlignment]::Center
        $format.FormatFlags = [Drawing.StringFormatFlags]::NoWrap
        $g.Clear([Drawing.Color]::Transparent)
        if ($name -eq 'dot') { $g.FillEllipse([Drawing.Brushes]::Black,12,2,12,12) }
        else { $g.DrawString('100%',$font,[Drawing.Brushes]::Black,[Drawing.RectangleF]::new(0,0,36,16),$format) }
        $entry = $zip.CreateEntry("customUI/$name.png").Open()
        try { $bitmap.Save($entry,[Drawing.Imaging.ImageFormat]::Png) }
        finally { $entry.Dispose(); $g.Dispose(); $bitmap.Dispose(); $font.Dispose(); $format.Dispose() }
    }
} finally { $zip.Dispose(); $file.Dispose() }
