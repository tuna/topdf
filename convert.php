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
    $old_level = error_reporting();
    // Ensure to report E_NOTICE.
    error_reporting(E_ALL);
    try {
        $ret = iconv('UTF-8', 'GBK', $str);
    } catch (Exception $e) {
        $ret = '';
    }
    error_reporting($old_level);
    return $ret;
}
// Assume no error when convert form GBK to UTF-8.
function strenc_fromlocal($str) {
    return iconv('GBK', 'UTF-8', $str);
}
// Use system locale settings.
// On Simplified Chinese Windows it's CP936.
$locale = setlocale(LC_ALL, '');
$filename_utf = $_FILES['upload']['name'];
$filename_tmp = $_FILES['upload']['tmp_name']; // path\to\phpXXX.tmp
// Convert string encoding of the filename.
$filename = strenc_tolocal($filename_utf);
$path_parts = pathinfo($filename);
$filename_ext = strtolower($path_parts['extension']);
$filename_name = $path_parts['filename']; // filename without extension
$src_filename_rel="./upload/$filename";
$cur_dir = getcwd();

// File type(ext) validation.
$accepted_file_types = array(
    'doc' => 'application/msword',
//    'ppt' => 'application/vnd.ms-powerpoint',
//    'xls' => 'application/ms-excel',
//    'docx' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    // Unfortunately 'file' currently cannot peek inside Office2007 archive.
    'docx' => 'application/zip',
//    'pptx' => 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
//    'xlsx' => 'application/vnd.openxmlformats-officedocument.spreahsheetml.sheet',
);
// NOTE: function mime_content_type is deprecated since PHP 5.3
// and it is undefined if PHP is not compiled with the option mime-magic
//var_dump(mime_content_type($filename_tmp));
// However, we might use file.exe by cygwin.
$file_bin = 'C:/cygwin/bin/file.exe';
if (file_exists($file_bin)) {
    $filename_tmp_posix = str_replace('C:', '/cygdrive/c',
        strtr($filename_tmp, "\\", '/'));
    $command_str = $file_bin . ' -b --mime-type ' .
         $filename_tmp_posix . ' 2>&1';
    #print '$command_str:'; var_dump($command_str);
    $file_mime_type = exec($command_str);
    print '$file_mime_type:'; var_dump($file_mime_type);
    $file_valid = (array_key_exists($filename_ext, $accepted_file_types)) &&
        ($accepted_file_types[$filename_ext] === $file_mime_type);
    #print '$filename_valid:'; var_dump($file_valid);
} else {
    // Do simple validation by file extension.
    $file_valid = (array_key_exists($filename_ext, $accepted_file_types));
}

if (! $file_valid) {
    print "<p>Invalid file.</p>" . "\n";
    exit;
}

if (move_uploaded_file($filename_tmp, $src_filename_rel))
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
    $link = 'download/' .
        rawurlencode(strenc_fromlocal($filename_name)) . '.pdf';
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
