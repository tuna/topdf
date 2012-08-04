<?php
// vim: set sw=4 sts=4 et:
/**
 * Convert uploaded file to pdf.
 */
$error_page_html_head = '<!doctype html>
<html>
    <head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>转换</title>
    </head>
<body>
';
$error_page_html_tail = '
</body>
</html>';

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
    #print '$file_mime_type:'; var_dump($file_mime_type);
    $file_valid = (array_key_exists($filename_ext, $accepted_file_types)) &&
        ($accepted_file_types[$filename_ext] === $file_mime_type);
    #print '$file_valid:'; var_dump($file_valid);
} else {
    // Do simple validation by file extension.
    $file_valid = (array_key_exists($filename_ext, $accepted_file_types));
}

if (! $file_valid) {
    echo $error_page_html_head;
    echo "<p>Invalid file.</p>" . "\n";
    echo $error_page_html_tail;
    exit;
}

if ($_FILES['upload']['error'] === UPLOAD_ERR_OK) {
    $src_filename = preg_replace('/tmp$/', $filename_ext, $filename_tmp);
    $pdf_filename = preg_replace('/tmp$/', 'pdf', $filename_tmp);
    #print 'src: ' . strenc_fromlocal($src_filename) . "<br />\n";
    #print 'pdf: ' . strenc_fromlocal($pdf_filename) . "<br />\n";

    $ret = move_uploaded_file($filename_tmp, $src_filename);
    if (! $ret) {
        // Handle move_uploaded_file failure.
        echo $error_page_html_head;
        echo '<p>噢不好了：移动上传文件失败</p>' . "\n";
        echo $error_page_html_tail;
    }
    try {
        $wps = new COM("WPS.Application");
        $doc = $wps->Documents->Open($src_filename);
        $doc->exportpdf($pdf_filename);
        $doc->Close();
        unset($doc, $wps);
        $fname = 'download.pdf';
        $fnamestar = rawurlencode(strenc_fromlocal($filename_name)) . '.pdf';
        // Set headers for downloading.
        header('Cache-Control: no-cache, must-revalidate');
        header("Expires: Sat, 26 Jul 1997 05:00:00 GMT"); // Date in the past
        header('Content-Type: application/pdf');
        header('Content-Disposition: attachment; filename=' . $fname .
            "; filename*=utf-8\\''" . $fnamestar);
        header('Content-Length: ' . filesize($pdf_filename));
        readfile($pdf_filename);
        unlink($pdf_filename);
        unlink($src_filename);
        exit;
    } catch (Exception $e) {
        echo $error_page_html_head;
        echo '<p>生成PDF失败 -__-</p>' . "\n";
        echo '<p>异常信息：' . $e->getMessage() . "</p>\n";
        echo '<pre>';
        echo $e->getTraceAsString() . "\n";
        echo '</pre>';
        echo $error_page_html_tail;
    }
} else {
    /*
     * TODO
    switch ($_FILES['upload']['error'])
    {
    }
     */
    echo $error_page_html_head;
    echo '<p>文件上传失败呃</p>' . "\n";
    echo $error_page_html_tail;
}
