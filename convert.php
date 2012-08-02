<?php
/**
 * Convert uploaded file to pdf.
 */
?><!doctype html>
<html>
    <head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>转换</title>
    </head>
<body>
<?php
function strenc_tolocal($str) {
    return iconv('UTF-8', 'GBK', $str);
}
function strenc_fromlocal($str) {
    return iconv('GBK', 'UTF-8', $str);
}
// Use system locale settings.
// On Simplified Chinese Windows it's CP936.
$locale = setlocale(LC_ALL, '');
print 'locale: ' . $locale . "\n";
$filename_utf = $_FILES['upload']['name'];
// Convert string encoding of the filename.
$filename = strenc_tolocal($filename_utf);
$path_parts = pathinfo($filename);
$filename_ext = $path_parts['extension'];
$filename_name = $path_parts['filename']; // filename without extension
$src_filename_rel="./upload/$filename";
$cur_dir = getcwd();
if (move_uploaded_file($_FILES['upload']['tmp_name'], $src_filename_rel))
{
    print "<p>上传成功</p>" . "\n";
    $wps = new COM("WPS.Application");
    $src_filename = $cur_dir . '/' . $src_filename_rel;
    $pdf_filename_rel = './download/' . $filename_name . '.pdf';
    $pdf_filename = $cur_dir . '/' . $pdf_filename_rel;
    print 'src: ' . strenc_fromlocal($src_filename) . "<br />\n";
    print 'pdf: ' . strenc_fromlocal($pdf_filename) . "<br />\n";
    $doc = $wps->Documents->Open($src_filename);
    $doc->exportpdf($pdf_filename);
    $doc->Close();
    unset( $doc , $wps );
    $link = strenc_fromlocal($pdf_filename_rel);
    print '<p><a href="' . $link . '">下载 PDF</a></p>' . "\n";
}
else
{
    switch ($_FILES['upload']['error'])
    {
    case 1:
        print '<p>The file is bigger than this PHP installation allows</p>';
        break;
    case 2:
        print '<p>The file is bigger than this form allows</p>';
        break;
    case 3:
        print '<p>Only part of the file was uploaded</p>';
        break;
    case 4:
        print '<p>No file was uploaded</p>';
        break;
    }
}
?>
</body>
</html>
<!-- vim: set sw=4 sts=4 et: -->
