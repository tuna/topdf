window.onload = function () {
    var sec = 3;
    var msg = sec + ' 秒后自动回到首页';
    document.body.innerHTML += '<p style="color: orange;">' + msg + '</p>';
    setTimeout(function () {
        window.location.href = 'index.html';
    }, sec * 1000);
}
