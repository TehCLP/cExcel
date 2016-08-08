
(function ($) {
    'use strict'

    var _prefix = $('#h_prefix').val();

    function getPath(url) {
        var _url = _prefix + url;
        return _url;
    }
    function getFileExtension(filename) {
        return (/[.]/.exec(filename)) ? '' + /[^.]+$/.exec(filename) : ''; //undefined;
    }
    function chkFormData() {
        var b = false;
        if (window.FormData !== undefined) {
            b = true;
        } else {
            alert("This browser doesn't support HTML5 file uploads!");
        }
        return b;
    }
    function upload(files, data, func) {

        var fileName = files[0].name;
        var ext = getFileExtension(fileName).toLowerCase();

        if (ext == 'xls' || ext == 'xlsx') {

            data.append('ext', ext);
            data.append(fileName, files[0]);

            $.ajax({
                type: "POST",
                url: getPath('excel/upload'),
                contentType: false,
                processData: false,
                data: data,
                success: func,
                error: function (xhr, status, p3, p4) {
                    var err = "Error " + " " + status + " " + p3 + " " + p4;
                    if (xhr.responseText && xhr.responseText[0] == "{")
                        err = JSON.parse(xhr.responseText).Message;
                    console.log(err);
                }
            });

        } else {
            alert("Upload File .xls, .xlsx only !!!");
        }

    }

    $('#btnUploadFile').on('click', function () {

        var files = $('#file')[0].files;
        if (files.length > 0) {

            if (confirm('ต้องการ Upload ???')) {

                if (chkFormData()) {

                    var data = new FormData();
                    data.append('fileName', files[0].name);

                    upload(files, data, function (result) {
                        var _res = result;
                        console.log(_res);
                        location.reload();
                    });
                }

            }
        }

    });

})(jQuery);