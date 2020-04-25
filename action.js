const electron = require('electron');
const { ipcRenderer } = electron;

function openFile() {
    ipcRenderer.send('crawl:openFile', true);
}


ipcRenderer.on("crawl:error_choose_not_txt", (e, item) => {
    if (item) {
        document.getElementById("error_crawl").innerHTML = "Tệp danh sách số điện thoại cần phải là tệp .txt";
        document.getElementById("error_crawl").style.display = 'block';
        document.getElementById("btn_crawl").style.display = 'flex';
        //document.getElementById("error_text").style.display = 'none';
        document.getElementById("span_file_input_error").style.display = 'block';
        document.getElementById("span_file_input_error").innerHTML = "Tệp danh sách số điện thoại cần phải là tệp .txt.Bám vào đây đẻ chọn lại";
        document.getElementById("span_file_input_success").style.display = 'none';
    } else {
        document.getElementById("div-login-file-input").style.display = 'flex';
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

ipcRenderer.on("crawl:result", (e, item) => {
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
        if (newFileNameTxt != ""){
            document.getElementById("span_file_input_success").innerHTML = "Truy xuất dữ liệu từ tệp " + fileNameTXT + " thành công.Tệp chuẩn bị là "+newFileNameTxt;
        }else {
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
ipcRenderer.on("crawl:read_sucess_first_time", (e, item) => {
    fileNameTXT = item;
    // document.getElementById("div_login_loading").style.display = 'block';
    // document.getElementById("div_progress_bar").style.display = 'block'; 
    //document.getElementById("error_crawl").style.display = 'none';
    document.getElementById("btn_crawl").style.display = 'flex';
    document.getElementById("div_delay_time").style.display = 'flex';
    document.getElementById("error_crawl").style.display = 'none';
    document.getElementById("span_file_input_success").style.display = 'block';
    document.getElementById("span_file_input_success").innerHTML = "Tệp bạn chọn tên là '" + fileNameTXT + "'.Bấm vào đây để chọn lại tệp";
    document.getElementById("span_file_input_error").style.display = 'none';

});

//đã chọn mới một ffile txt khác
ipcRenderer.on("crawl:read_sucess_new", (e, item) => {
    newFileNameTxt = item ;
    // document.getElementById("div_login_loading").style.display = 'block';
    // document.getElementById("div_progress_bar").style.display = 'block'; 
    //document.getElementById("error_crawl").style.display = 'none';
    document.getElementById("btn_crawl").style.display = 'flex';
    document.getElementById("div_delay_time").style.display = 'flex';
    document.getElementById("error_crawl").style.display = 'none';
    document.getElementById("span_file_input_success").style.display = 'block';
    document.getElementById("span_file_input_success").innerHTML = "Bạn mới chọn một tệp mới là '" + newFileNameTxt + "'.Bấm vào đây để chọn lại tệp";
    document.getElementById("span_file_input_error").style.display = 'none';

});

ipcRenderer.on("crawl:read_error", (e, item) => {
    newFileNameTxt = "";
    // document.getElementById("div_login_loading").style.display = 'none';
    // document.getElementById("div_progress_bar").style.display = 'none';
    // document.getElementById("success_text").style.display = 'none';
    //document.getElementById("error_crawl").innerHTML = `Lỗi đọc file,cần chọn file txt,bấm tổ hợp crt + f để chọn file`;
    document.getElementById("span_file_input_error").style.display = 'block';
    document.getElementById("span_file_input_error").innerHTML = "Tệp hiện tại '" + item + "'hiện không đọc được, vui lòng bấm vào đây chọn lại tệp";
    document.getElementById("span_file_input_success").style.display = 'none';
    //document.getElementById("error_crawl").style.display = 'block';
    //crawling = false;

});

ipcRenderer.on("crawl:read_error_null", (e, item) => {
    newFileNameTxt = "";
    // document.getElementById("div_login_loading").style.display = 'none';
    // document.getElementById("div_progress_bar").style.display = 'none';
    // document.getElementById("success_text").style.display = 'none';
    //document.getElementById("error_crawl").innerHTML = "File "+fileNameTXT+" chưa có số điện thoại nào";
    document.getElementById("span_file_input_error").style.display = 'block';
    document.getElementById("span_file_input_error").innerHTML = "Tệp '" + item + "' chưa có số điện thoại nào,bấm vào đây để chọn lại tệp";
    document.getElementById("span_file_input_success").style.display = 'none';
    document.getElementById("error_crawl").style.display = 'block';
    // crawling = false;

});

ipcRenderer.on("crawl:network_error", (e, item) => {
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

ipcRenderer.on("crawl:incorrect_number", (e, item) => {
    console.log(item);
    if (item) {
        // document.getElementById("div_login_loading").style.display = 'none';
        // document.getElementById("div_progress_bar").style.display = 'none';
        // document.getElementById("success_text").style.display = 'none';
        document.getElementById("error_crawl").innerHTML = "Số điện thoại '" + item + "'  không đúng! Chương trình sẽ không tra cứu số điện thoại này";
        document.getElementById("error_crawl").style.display = 'block';
    }
});

ipcRenderer.on("crawl:onrunning", (e, item) => {
    document.getElementById("error_crawl").style.display = 'none';
    let tItem = item.split(" ");
    let tResult = Math.round(Number.parseFloat(tItem[0]) / Number.parseFloat(tItem[1]) * 100 * 100) / 100;
    document.getElementById("div_grey").style.width = tResult + "%";
    document.getElementById("success_text").innerHTML = "Tệp ''"+fileNameTXT+"'' --- Đã hoàn thành " + tResult + "% - ( " + tItem[0] + "/" + tItem[1] + " )";

});

ipcRenderer.on("crawl:runwithfile", (e, item) => {
    fileNameTXT = item;
    document.getElementById("span_file_input").innerHTML = "Đang tra cứu danh sách số trong tệp '" + fileNameTXT + "'.Bấm vào đây để đổi lại tệp";
});

ipcRenderer.on("crawl:hideBTN", (e, item) => {
    if (item) {
        document.getElementById("btn_crawl").style.display = 'none';
        document.getElementById("div_delay_time").style.display = 'none';
    }
});

ipcRenderer.on("crawl:login_success", (e, item) => {
    if (item === 1) {
        loginSuccess();
    } else if (item === 0){
        document.getElementById("error_text").innerHTML = "Sai tên đăng nhập hoặc mật khẩu";
        document.getElementById("error_text").style.color = 'red';
        document.getElementById("error_text").style.display = 'block';
        //document.getElementById("error_text").classList.remove('logining');
    }else if (item === -1){
        document.getElementById("error_text").innerHTML = "Có lỗi khi đăng nhập,vui lòng thử lại";
        document.getElementById("error_text").style.color = 'red';
        document.getElementById("error_text").style.display = 'block';
       // document.getElementById("error_text").classList.remove('logining');
    } else if (item == 2) {
        document.getElementById("error_text").innerHTML = "Đang đăng nhập vui lòng đợi ....";
        document.getElementById("error_text").style.color = 'green';
        document.getElementById("error_text").style.display = 'block';
       // document.getElementById("error_text").classList.add('logining');
    }
});

function login(){

}

function loginSuccess() {
    document.getElementById("div_login").style.display = 'none';
    document.getElementById("error_text").style.display = 'none';
    document.getElementById("div-login-success").style.display = 'block';
    setTimeout(() => {
        document.getElementById("div-craw").style.display = 'block';
        document.getElementById("div-login-success").style.display = 'none';
    }, 800)
}

function otp(){

}

function crawl(){

}

function openFile(){
    ipcRenderer.send('crawl:openFile', true);
}