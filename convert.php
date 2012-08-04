<?php
// vim: set sw=4 sts=4 et:
/**
 * Convert uploaded file to pdf.
 */
$error_page_html_head = '<!doctype html>
<html>
    <head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link href="reset.css" rel="stylesheet" type="text/css" />
        <link href="errorpage.css" rel="stylesheet" type="text/css" />
        <script type="text/javascript" src="back.js"></script>
        <title>转换错误</title>
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

// Make empty upload an error.
// Credit goes to
// http://www.php.net/manual/en/features.file-upload.errors.php#107295
define("UPLOAD_ERR_EMPTY", 5);
if($_FILES['upload']['size'] == 0 && $_FILES['upload']['error'] == 0) {
    $_FILES['upload']['error'] = 5;
}
$upload_errors = array(
    UPLOAD_ERR_OK         => "上传成功。",
    UPLOAD_ERR_INI_SIZE   => "文件过大，超过 upload_max_filesize.",
    UPLOAD_ERR_FORM_SIZE  => "文件过大，超过 MAX_FILE_SIZE.",
    UPLOAD_ERR_PARTIAL    => "文件仅有部分上传。",
    UPLOAD_ERR_NO_FILE    => "没有文件上传。",
    UPLOAD_ERR_NO_TMP_DIR => "缺少临时文件夹。",
    UPLOAD_ERR_CANT_WRITE => "无法写入磁盘。",
    UPLOAD_ERR_EXTENSION  => "上传被 PHP 扩展阻止。",
    UPLOAD_ERR_EMPTY      => "文件内容为空。" // add this to avoid an offset
);

$err = $_FILES['upload']['error'];
if ($err !== UPLOAD_ERR_OK) {
    echo $error_page_html_head;
    echo '<p>文件上传失败呃</p>' . "\n";
    echo '<p>具体原因：' . $upload_errors[$err] . '</p>' . "\n";
    echo $error_page_html_tail;
    exit;
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
error_log('$filename_tmp: ' . $filename_tmp);
error_log('rawurlencoded filename: '. rawurlencode($filename_utf));
error_log('$filename_ext: ' . $filename_ext);

// File type(ext) validation.
$accepted_file_types = array(
    'wps' => array('application/msword'),
    'doc' => array('application/msword'),
    'ppt' => array('application/vnd.ms-powerpoint', 'application/msword'),
    'xls' => array('application/ms-excel', 'application/msword'),
    'docx' => array('application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/zip'),
    'pptx' => array('application/vnd.openxmlformats-officedocument.presentationml.presentation', 'application/zip'),
    'xlsx' => array('application/vnd.openxmlformats-officedocument.spreahsheetml.sheet', 'application/zip'),
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
    $file_mime_type = exec($command_str);
} else {
    error_log('*** Trust mime type info by client side.');
    $file_mime_type = $_FILES['upload']['type'];
}

error_log('$file_mime_type: ' . $file_mime_type);
$file_valid = true;
$file_invalid_msg = '';
if (array_key_exists($filename_ext, $accepted_file_types)) {
    if (in_array($file_mime_type, $accepted_file_types[$filename_ext], true)) {
        // Valid file. wow
    } else {
        $file_valid = false;
        $file_invalid_msg = '错误的 MIME 类型：' . $file_mime_type;
    }
} else {
    $file_valid = false;
    $file_invalid_msg = '不支持的文件格式：' . $filename_ext;
}
if (! $file_valid) {
    echo $error_page_html_head;
    echo "<p>文件格式错误：</p>" . "\n";
    echo "<p>" . $file_invalid_msg . "</p>\n";
    echo $error_page_html_tail;
    exit;
}

$src_filename = preg_replace('/tmp$/', $filename_ext, $filename_tmp);
$pdf_filename = preg_replace('/tmp$/', 'pdf', $filename_tmp);
error_log('src: ' . strenc_fromlocal($src_filename));
error_log('pdf: ' . strenc_fromlocal($pdf_filename));

// Rename tmp file to have the right extension.
$ret = move_uploaded_file($filename_tmp, $src_filename);
if (! $ret) {
    // Handle move_uploaded_file failure.
    echo $error_page_html_head;
    echo '<p>噢不好了：移动上传文件失败</p>' . "\n";
    echo $error_page_html_tail;
}
// Really do conversion.
try {
    switch ($filename_ext) {
    case 'doc':
    case 'docx':
    case 'wps':
        $wps = new COM("WPS.Application");
        $doc = $wps->Documents->Open($src_filename);
        break;
    case 'ppt':
    case 'pptx':
        $wps = new COM("WPP.Application");
        $doc = $wps->Presentations->Open($src_filename);
        break;
    case 'xls':
    case 'xlsx':
        $wps = new COM("ET.Application");
        $doc = $wps->Workbooks->Open($src_filename);
        break;
    default:
        error_log('OMG, you should not come here!');
        echo $error_page_html_head;
        echo '<p>你穿越了！赶快回去吧。</p>' . "\n";
        echo $error_page_html_tail;
        unlink($src_filename);
        exit;
        break;
    }
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
    unlink($src_filename);
    exit;
}
