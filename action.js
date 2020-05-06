const electron = require('electron');
const { ipcRenderer } = electron;

var crawling = false;

var fileNameTXT = "";
var newFileNameTxt = "";
var currentOTPStatus = -100;

const crawlCommand = {
    login: "crawl:login",
    otp: "crawl:otp",
    openFile: "crawl:openFile",
    wrongPhoneNumber: "crawl:incorrect_number",
    hideBTN: "crawl:hideBTN",
    networkError: "crawl:network_error",
    result: "crawl:result",
    readError: "crawl:read_error",
    readErrorNull: "crawl:read_error_null",
    readSuccess: "crawl:read_sucess_new",
    readSuccessFirtTime: "crawl:read_sucess_first_time",
    inputfileNotexcel: "crawl:error_choose_not_xlsx",
    doCrawl: "crawl:do",
    readrSuccessNew: "crawl:read_sucess_new",
    runWithFile: "crawl:runwithfile",
    onRunning: "crawl:onrunning",
    loginSuccess: "crawl:login_success",
    log: "crawl:log",
    allowToWrite:"crawl:log", // cho phép write hoặc không, mắc định là cho phép, chỉ khi có dialog , mất kết nối mạng, hoặc sesion timeout , số không hợp lệ, không tìm thấy số
};

function openFile() {
    ipcRenderer.send(crawlCommand.openFile, true);
}

ipcRenderer.on(crawlCommand.log, (e, item) => {
    console.log("puppeteer log", item);
});

ipcRenderer.on(crawlCommand.inputfileNotexcel, (e, item) => {
    if (item) {
        document.getElementById("error_crawl").innerHTML = "Tệp danh sách số điện thoại cần phải là tệp .xlsx";
        document.getElementById("error_crawl").style.display = 'block';
        document.getElementById("btn_crawl").style.display = 'flex';
        //document.getElementById("error_text").style.display = 'none';
        document.getElementById("span_file_input_error").style.display = 'block';
        document.getElementById("span_file_input_error").innerHTML = "Tệp danh sách số điện thoại cần phải là tệp .xlsx.Bấm vào đây đẻ chọn lại";
        document.getElementById("span_file_input_success").style.display = 'none';
    } else {
        document.getElementById("crawl_login_file_input").style.display = 'flex';
    }
});

ipcRenderer.on("crawl:error_choose_not_chrome", (e, item) => {
    if (item) {
        document.getElementById("error_crawl").innerHTML = "File google chrome phải là file exe";
        document.getElementById("error_crawl").style.display = 'block';
        document.getElementById("error_text").innerHTML = "File google chrome phải la file exe";
        document.getElementById("error_text").style.display = 'block';
    } else {
    }
});

ipcRenderer.on(crawlCommand.result, (e, item) => {
    if (item) {
        document.getElementById("div_login_loading").style.display = 'none';
        document.getElementById("div_progress_bar").style.display = 'none';
        document.getElementById("success_text").style.display = 'none';
        document.getElementById("success_text").innerHTML = "Truy xuất dữ liệu thành công";
        document.getElementById("success_text").style.display = 'block';
        document.getElementById("error_crawl").style.display = 'none';
        document.getElementById("btn_crawl").style.display = 'flex';
        document.getElementById("div_delay_time").style.display = 'flex';

        document.getElementById("span_file_input_error").style.display = 'none';
        if (newFileNameTxt != "") {
            document.getElementById("span_file_input_success").innerHTML = "Truy xuất dữ liệu từ tệp " + fileNameTXT + " thành công.Tệp chuẩn bị là " + newFileNameTxt;
        } else {
            document.getElementById("span_file_input_success").innerHTML = "Truy xuất dữ liệu từ tệp " + fileNameTXT + " thành công.Bấm vào đây để chọn lại tệp";
        }
        document.getElementById("span_file_input_success").style.display = 'block';
        //crawling = false;
    } else {
        document.getElementById("div_login_loading").style.display = 'none';
        document.getElementById("div_progress_bar").style.display = 'none';
        document.getElementById("success_text").style.display = 'none';
        document.getElementById("error_crawl").innerHTML = "Truy xuất dữ liệu không thành công";
        document.getElementById("error_crawl").style.display = 'block';
        document.getElementById("btn_crawl").style.display = 'flex';
        document.getElementById("div_delay_time").style.display = 'flex';
        document.getElementById("span_file_input_error").style.display = 'block';
        document.getElementById("span_file_input_error").innerHTML = "Truy xuất dữ liệu từ tệp " + fileNameTXT + " thành công.Bấm vào đây để chọn lại tệp";
        document.getElementById("span_file_input_success").style.display = 'none';
        //crawling = false;
    }
});

//lần đầu chạy
ipcRenderer.on(crawlCommand.readSuccessFirtTime, (e, item) => {
    fileNameTXT = item;
    document.getElementById("btn_crawl").style.display = 'flex';
    document.getElementById("div_delay_time").style.display = 'flex';
    document.getElementById("error_crawl").style.display = 'none';
    document.getElementById("span_file_input_success").style.display = 'block';
    document.getElementById("span_file_input_success").innerHTML = "Tệp bạn chọn tên là '" + fileNameTXT + "'.Bấm vào đây để chọn lại tệp";
    document.getElementById("span_file_input_error").style.display = 'none';

});

//đã chọn mới một file txt khác
ipcRenderer.on(crawlCommand.readrSuccessNew, (e, item) => {
    newFileNameTxt = item;
    document.getElementById("btn_crawl").style.display = 'flex';
    document.getElementById("div_delay_time").style.display = 'flex';
    document.getElementById("error_crawl").style.display = 'none';
    document.getElementById("span_file_input_success").style.display = 'block';
    document.getElementById("span_file_input_success").innerHTML = "Bạn mới chọn một tệp mới là '" + newFileNameTxt + "'.Bấm vào đây để chọn lại tệp";
    document.getElementById("span_file_input_error").style.display = 'none';

});

ipcRenderer.on(crawlCommand.readError, (e, item) => {
    newFileNameTxt = "";
    document.getElementById("span_file_input_error").style.display = 'block';
    document.getElementById("span_file_input_error").innerHTML = "Tệp hiện tại '" + item + "'hiện không đọc được, vui lòng bấm vào đây chọn lại tệp";
    document.getElementById("span_file_input_success").style.display = 'none';
    //document.getElementById("error_crawl").style.display = 'block';
    //crawling = false;

});

ipcRenderer.on(crawlCommand.readErrorNull, (e, item) => {
    newFileNameTxt = "";
    document.getElementById("span_file_input_error").style.display = 'block';
    document.getElementById("span_file_input_error").innerHTML = "Tệp '" + item + "' chưa có số điện thoại nào,bấm vào đây để chọn lại tệp";
    document.getElementById("span_file_input_success").style.display = 'none';
    document.getElementById("error_crawl").style.display = 'block';
    // crawling = false;

});

ipcRenderer.on(crawlCommand.networkError, (e, item) => {
    if (item) {
        document.getElementById("div_login_loading").style.display = 'none';
        document.getElementById("div_progress_bar").style.display = 'none';
        document.getElementById("success_text").style.display = 'none';
        document.getElementById("error_crawl").innerHTML = "Lỗi mạng,vui lòng thử lại";
        document.getElementById("error_crawl").style.display = 'block';
        document.getElementById("btn_crawl").style.display = 'flex';
        //crawling = false;
    } else {
        document.getElementById("div_login_loading").style.display = 'none';
        document.getElementById("div_progress_bar").style.display = 'none';
        document.getElementById("success_text").style.display = 'none';
        document.getElementById("error_crawl").innerHTML = "Lỗi mạng,vui lòng thử lại";
        document.getElementById("error_crawl").style.display = 'block';
        document.getElementById("btn_crawl").style.display = 'flex';
        //crawling = false;
    }
    document.getElementById("span_file_input").innerHTML = "Đang tra cứu danh sách số trong tệp '" + fileNameTXT + "'(Đang lỗi mạng)Bấm vào đây để đổi lại tệp";
});

ipcRenderer.on(crawlCommand.wrongPhoneNumber, (e, item) => {
    console.log(item);
    if (item) {
        document.getElementById("error_crawl").innerHTML = "Số điện thoại '" + item + "'  không đúng! Chương trình sẽ không tra cứu số điện thoại này";
        document.getElementById("error_crawl").style.display = 'block';
    }
});

ipcRenderer.on(crawlCommand.onRunning, (e, item) => {
    document.getElementById("error_crawl").style.display = 'none';
    let tItem = item.split(" ");
    let tResult = Math.round(Number.parseFloat(tItem[0]) / Number.parseFloat(tItem[1]) * 100 * 100) / 100;
    document.getElementById("div_grey").style.width = tResult + "%";
    document.getElementById("success_text").innerHTML = "Tệp ''" + fileNameTXT + "'' --- Đã hoàn thành " + tResult + "% - ( " + tItem[0] + "/" + tItem[1] + " )";

});

ipcRenderer.on(crawlCommand.runWithFile, (e, item) => {
    fileNameTXT = item;
    document.getElementById("span_file_input").innerHTML = "Đang tra cứu danh sách số trong tệp '" + fileNameTXT + "'.Bấm vào đây để đổi lại tệp";
});

ipcRenderer.on(crawlCommand.hideBTN, (e, item) => {
    if (item) {
        document.getElementById("btn_crawl").style.display = 'none';
        document.getElementById("div_delay_time").style.display = 'none';
    }
});

ipcRenderer.on(crawlCommand.loginSuccess, (e, item) => {
    console.log("loginSuccess", item);
    hideProgressBarLogin();
    if (item === 1) {
        loginSuccess();
    } else if (item === 0) {
        document.getElementById("crawl_login_error_text").innerHTML = "Sai tên đăng nhập hoặc mật khẩu";
        document.getElementById("crawl_login_error_text").style.color = 'red';
        document.getElementById("crawl_login_error_text").style.display = 'block';
    } else if (item === -1) {
        let tempValue = document.getElementById("crawl_login_error_text").innerHTML;
        if (tempValue == "Đang đăng nhập vui lòng đợi .... ||" || tempValue == null) {
            document.getElementById("crawl_login_error_text").innerHTML = "Có lỗi khi đăng nhập,vui lòng thử lại";
            document.getElementById("crawl_login_error_text").style.color = 'red';
            document.getElementById("crawl_login_error_text").style.display = 'block';
        }
    } else if (item == 2) {
        showProgressBarLogin()
        document.getElementById("crawl_login_error_text").innerHTML = "Đang đăng nhập vui lòng đợi ....";
        document.getElementById("crawl_login_error_text").style.color = 'green';
        document.getElementById("crawl_login_error_text").style.display = 'block';
    } else if (item == -2) {
        document.getElementById("crawl_login_error_text").innerHTML = "Mật khẩu phải ít nhất 8 ký tự, 1 ký tự hoa, 1 ký tự đặc biệt, 1 ký tự số";
        document.getElementById("crawl_login_error_text").style.color = 'red';
        document.getElementById("crawl_login_error_text").style.display = 'block';
    } else if (item == -3) {
        document.getElementById("crawl_login_error_text").innerHTML = "Sai mật khẩu ....";
        document.getElementById("crawl_login_error_text").style.color = 'red';
        document.getElementById("crawl_login_error_text").style.display = 'block';
    }
});


function login() {
    document.getElementById("crawl_login_error_text").style.display = 'none';
    let user = document.getElementById("username").value;
    let pass = document.getElementById("password").value;

    if (!user) {
        document.getElementById("username").focus();
        document.getElementById("crawl_login_error_text").innerHTML = "Cần nhập tên đăng nhập ";
        document.getElementById("crawl_login_error_text").style.display = 'block';
    } else if (!pass) {
        document.getElementById("password").focus();
        document.getElementById("crawl_login_error_text").innerHTML = "Cần nhập mật khẩu";
        document.getElementById("crawl_login_error_text").style.display = 'block';
    } else {
        showProgressBarLogin();
        ipcRenderer.send(crawlCommand.login, user + " " + pass);
    }
}

function showProgressBarLogin() {
    document.getElementById("crawl_login_button_submit").style.display = 'none';
    document.getElementById("crawl_login_progress_bar").style.display = 'block';
}

function hideProgressBarLogin() {
    document.getElementById("crawl_login_button_submit").style.display = 'flex';
    document.getElementById("crawl_login_progress_bar").style.display = 'none';
}

function loginSuccess() {
    document.getElementById("crawl_login").style.display = 'none';
    document.getElementById("crawl_login_error_text").style.display = 'none';
    //màn hình chuyển từ login qua otp
    document.getElementById("crawl_login_success").style.display = 'flex';
    setTimeout(() => {
        //hiện otp
        document.getElementById("crawl_otp").style.display = 'flex';
        document.getElementById("crawl_otp_error_text").style.display = 'none';
        document.getElementById("crawl_login_success").style.display = 'none';
    }, 850)
}

ipcRenderer.on(crawlCommand.otp, (e, item) => {
    console.log("otp", item);
    hideProgressBarOTP();
    if (item === 1) {
        if (currentOTPStatus == -100) {//nếu như currentOTPSTAUS đang không phải là -100
            otpSuccess();
        }
    } else if (item === 0) {
        currentOTPStatus = item;
        document.getElementById("crawl_otp_error_text").innerHTML = "Mật khẩu OTP không đúng, vui lòng điền lại";
        document.getElementById("crawl_otp_error_text").style.color = 'red';
        document.getElementById("crawl_otp_error_text").style.display = 'block';
    } else if (item === -1) {
        currentOTPStatus = item;
        document.getElementById("crawl_otp_error_text").innerHTML = "Có lỗi khi xác thực OTP, vui lòng thử lại";
        document.getElementById("crawl_otp_error_text").style.color = 'red';
        document.getElementById("crawl_otp_error_text").style.display = 'block';
    } else if (item == 2) {
        showProgressBarOTP();
        document.getElementById("crawl_otp_error_text").innerHTML = "Đang xác thực OTP vui lòng đợi ....";
        document.getElementById("crawl_otp_error_text").style.color = 'green';
        document.getElementById("crawl_otp_error_text").style.display = 'block';
    } else if (item == -2) {
        currentOTPStatus = item;
        document.getElementById("crawl_otp_error_text").innerHTML = "Mật khẩu OTP đã hết hạn, bạn vui lòng đăng nhập lại ....";
        document.getElementById("crawl_otp_error_text").style.color = 'red';
        document.getElementById("crawl_otp_error_text").style.display = 'block';
        otpTimeOut();
    }
});


function otp() {
    document.getElementById("crawl_otp_error_text").style.display = 'none';
    showProgressBarOTP();
    currentOTPStatus = -100;
    let otpText = document.getElementById("otpText").value; console.log("otp", otpText);
    if (!otpText) {
        document.getElementById("otpText").focus();
        document.getElementById("crawl_otp_error_text").innerHTML = "Cần nhập mật khẩu OTP";
        document.getElementById("crawl_otp_error_text").style.display = 'block';
    } else {
        ipcRenderer.send(crawlCommand.otp, otpText);
    }
}

function showProgressBarOTP() {
    document.getElementById("crawl_otp_button_submit").style.display = 'none';
    document.getElementById("crawl_otp_progress_bar").style.display = 'block';
}

function hideProgressBarOTP() {
    document.getElementById("crawl_otp_button_submit").style.display = 'flex';
    document.getElementById("crawl_otp_progress_bar").style.display = 'none';
}

function otpSuccess() {
    document.getElementById("crawl_otp").style.display = 'none';
    document.getElementById("crawl_otp_error_text").style.display = 'none';
    //màn hình chuyển từ otp qua crawl
    document.getElementById("crawl_otp_success").style.display = 'flex';
    document.getElementById("crawl_otp_validating_text").style.display = "block";
    document.getElementById("crawl_otp_validating_text").style.color = "green";
    document.getElementById("crawl_otp_validating_text").style.color = "Xác thực mật khẩu OTP thành công";

    setTimeout(() => {
        //hiện crawl
        document.getElementById("div_craw").style.display = 'flex';
        document.getElementById("crawl_otp_success").style.display = 'none';
        document.getElementById("crawl_otp_validating_text").style.display = "none";
    }, 850)
}

function otpTimeOut() {
    document.getElementById("crawl_otp").style.display = 'none';
    document.getElementById("crawl_otp_error_text").style.display = 'none';
    //màn hình chuyển từ otp qua login
    document.getElementById("crawl_otp_success").style.display = 'flex';
    document.getElementById("crawl_otp_validating_text").style.display = "block";
    document.getElementById("crawl_otp_validating_text").style.color = "red";
    document.getElementById("crawl_otp_validating_text").style.color = "Mật khẩu OTP đã hết hạn, bạn vui lòng đăng nhập lại ....";

    setTimeout(() => {
        //hiện lại login
        document.getElementById("crawl_login").style.display = 'flex';
        document.getElementById("crawl_login_error_text").style.display = 'none';

        document.getElementById("crawl_otp_success").style.display = 'none';
        document.getElementById("crawl_otp_validating_text").style.display = "none";
    }, 850)
}

function crawl() {
    document.getElementById("error_crawl").style.display = 'none';
    document.getElementById("success_text").style.display = 'none';

    document.getElementById("div_login_loading").style.display = 'block';
    document.getElementById("div_progress_bar").style.display = 'block';

    //document.getElementById("crawl_login_file_input").style.display = 'flex';

    document.getElementById("success_text").style.display = 'block';
    document.getElementById("span_file_input_error").style.display = 'none';
    document.getElementById("success_text").innerHTML = "0%";
    document.getElementById("div_grey").style.width = "0%";
    let delayTime = document.getElementById("second_crawl").value;
    delayTime = delayTime * 1000;
    ipcRenderer.send(crawlCommand.doCrawl, delayTime);
    document.getElementById("btn_crawl").style.display = 'none';
    document.getElementById("div_delay_time").style.display = 'none';
    if (newFileNameTxt != "") {
        fileNameTXT = newFileNameTxt;
    }
    newFileNameTxt = "";
    document.getElementById("span_file_input_success").innerHTML = "Tệp bạn chọn tên là '" + fileNameTXT + "'.Bấm vào đây để chọn lại tệp";
    document.getElementById("span_file_input_success").style.display = 'block';
}

function openFile() {
    ipcRenderer.send(crawlCommand.openFile, true);
}