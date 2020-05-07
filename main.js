const electron = require('electron');
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
var xl = require('excel4node');

const { app, BrowserWindow, ipcMain, Menu, dialog } = electron;

var concurentPupl, concurentLogin;
var delayInMilliseconds = 10000;
var defaultDelay = 10000;
var inputPhoneNumberArray = [];
let fileNametxt = "";
var wb;
var ws;
var defaultHeight = 15;
var username = "";
var password = "";
//danh sách số điện thoại
let tResult = [];

let mainWindow;
var mainBrowser = null;
var exPath = '';
var startStartIndex = 0;
var rowSpacing = 2;
var directionToSource = "";
var wrongPassword = "Mật khẩu phải ít nhất 8 ký tự, 1 ký tự hoa, 1 ký tự đặc biệt, 1 ký tự số. Nếu không hợp lệ vui lòng đổi trước khi đăng nhập";
var lackPassword = "Mật khẩu";
var wrongLogin = "Tài khoản không hợp lệ, vui lòng thử lại";

var wrongOTP = "Mã OTP không hợp lệ, vui lòng thử lại";
var wrongPhoneNumber = "Số không hợp lệ";
var timoutOTP = "Phiên kiểm tra Otp hết hạn, hệ thống trở về trang đăng nhập.... ";
var headeTitle = "header";

var crawlUrl = "http://10.149.34.250:1609/Views/KhachHang/ThongTinKhachHang.aspx"; // vì nếu chưa dăng nhập thì vào trang lấy thông tin khách hàng cũng sẽ bị redirect về trang đăng nhập
var threshHoldeCount = 5;
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
    runWithFile: "crawl:runwithfile",
    onRunning: "crawl:onrunning",
    loginSuccess: "crawl:login_success",
    log: "crawl:log",
    signalWrite: "crawl:log", // cho phép write hoặc không, mắc định là cho phép, chỉ khi có dialog , 
    //số không hợp lệ, -1 
    // không tìm thấy số -2
    // hoặc sesion timeout , -3
    //mất kết nối mạng -4 - trường hợp ít xảy ra, khồng xét
};

var canWrite = true;
var mCheckTrue = "Mở", mCheckFalse = "Đóng";

var currentData = [
    // "STT":"",
    // "Số thuê bao":"",
    // "MSIN":"",
    // "Loại thuê bao":"",
    // "Gọi đi":"",
    // "Gọi đến":"",
    // "Loại SIM":"",
    // "Hạng hội viên":"",
    // "Tỉnh":"",
    // "Ngày KH":"",
    // "Mã KH":"",
    // "Mã CQ",
    // "Tên thuê bao":"",
    // "Ngày sinh":"",
    // "Số GT":"",
    // "Ngày cấp",
    // "Số PIN/PUK",
    // "Số PIN2/PUK2",
    // "Đối tượng",
    // "Địa chỉ chứng từ",
    // "Địa chỉ thanh toán",
    // "Địa chỉ thường trú",
    // "Tài khoản chính",
    // "Hạn sử dụng",
    // "Thuê bao trả trước được tham gia khuyến mại",
    // "Gói cước trả trước ưu tiên mời KH đăng ký",
    // //dịch vụ 3G
    // //dịch vụ 1
    // "service":[]
]

var page, pageLogin;
var breakPerSerrvice = 6;//có 6 cột dịch vụ
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800, height: 600, webPreferences: {
            nodeIntegration: true // dung được require trên html
        }
    });

    //dev tool
    mainWindow.webContents.openDevTools();

    mainWindow.on('crashed', () => {
        win.destroy();
        createWindow();
    });

    mainWindow.loadURL(`file://${__dirname}/index.html`);
    //mainWindow.loadURL(isDev ? 'http://localhost:3000' : `file://${__dirname}/../build/index.html`);

    // Build menu from template
    const mainMenu = Menu.buildFromTemplate(mainMenuTemplate);
    // Insert menu
    Menu.setApplicationMenu(mainMenu);

    mainWindow.on('closed', function () {
        mainWindow = null;
    })

    //callback array
    Object.defineProperty(currentData, "push", {
        configurable : true,
            enumerable : false,
            writable : true,
            value :async function(...args){
                let result = Array.prototype.push.apply(this,args);
                await mainWindow.webContents.send(crawlCommand.log, "push to array value "+result);
                return result;
            }
    });

}

//hàm nothing
function nothing() {

}

// Create menu template
const mainMenuTemplate = [
    {
        label: 'Chức năng',
        submenu: [
            {
                label: 'Chọn tệp chứa danh sách điện thoại',
                accelerator: process.platform == 'darwin' ? 'Command+F' : 'Ctrl+F',
                click() {
                    // if (crawling == false) {
                    chooseSource(readFile, nothing);
                    // }
                }
            },
            {
                label: 'Thoát',
                accelerator: process.platform == 'darwin' ? 'Command+Q' : 'Ctrl+Q',
                click() {
                    app.quit();
                }
            }
        ]
    }
];

app.on('ready', createWindow);

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit();
    }
})

app.on('activate', function () {
    if (mainWindow === null) {
        createWindow();
    }
})

//bấm vào menu để mở file
//dùng tại 3 chỗ,
// 1 khi bấm vào menu -> đọc file , chỉ đọc file không làm gì cả
// 2 khi bấm vào nút chọn file khác , chỉ đọc file không làm gì cả
// 3 trường hợp chưa chọn file mà bấm vào nút lấy dữ liệu để crawl , mở đọc file , đọc xong rồi crawl
async function chooseSource(callback1, callback2) {
    dialog.showOpenDialog({
        title: "Chọn đường dẫn tới file text chứa danh sách số điện thoại",
        properties: ['openFile', 'multiSelections']
    }, function (files) {
        if (files !== undefined) {
            // handle files
        }
    }).then(async (result) => {
        if (!result.filePaths[0].endsWith(".xlsx")) {
            await mainWindow.webContents.send(crawlCommand.inputfileNotexcel, true);
        } else {
            directionToSource = result.filePaths[0];
            await mainWindow.webContents.send(crawlCommand.inputfileNotexcel, false);
            callback1(callback2);
        }
    }).catch(err => {
        ////console.log(err);
    });
};

//chuẩn bị file excel
async function prepareExxcel(callback) {

    //khởi tạo mảng
    inputPhoneNumberArray = [];
    tResult.forEach(element => {
        inputPhoneNumberArray.push(element);
    });

    await mainWindow.webContents.send(crawlCommand.log, 'ghi dữ liệu từ excel vào bộ nhớ tiến hành crawl  ' + inputPhoneNumberArray);

    cTotal = inputPhoneNumberArray.length;

    let cTimee = new Date();

    wb = new xl.Workbook();
    ws = wb.addWorksheet(cTimee.getHours() + "-" + cTimee.getMinutes() + " " + cTimee.getDate() + "/" + (cTimee.getMonth() + 1) + "/" + cTimee.getFullYear());

    ws.column(1).setWidth(5);//STT
    ws.column(2).setWidth(15);//Số thuê bao
    ws.column(3).setWidth(15);//MSIN
    ws.column(4).setWidth(15);//Loại thuê bao
    ws.column(5).setWidth(7);//Gọi đi
    ws.column(6).setWidth(7);//Gọi đến
    ws.column(7).setWidth(11);//Loại SIM
    ws.column(8).setWidth(15);//Hạng hội viên
    ws.column(9).setWidth(5);//Tỉnh
    ws.column(10).setWidth(15);//Ngày KH
    ws.column(11).setWidth(10);//Mã KH
    ws.column(12).setWidth(10);//Mã CQ
    ws.column(13).setWidth(30);//Tên thuê bao
    ws.column(14).setWidth(15);//Ngày sinh
    ws.column(15).setWidth(15);//Số GT
    ws.column(16).setWidth(15);//Ngày cấp
    ws.column(17).setWidth(17);//Số PIN/PUK
    ws.column(18).setWidth(17);//Số PIN2/PUK2
    ws.column(19).setWidth(16);//Đối tượng
    ws.column(20).setWidth(20);//Địa chỉ chứng từ
    ws.column(21).setWidth(25);//Địa chỉ thanh toán
    ws.column(22).setWidth(65);//Địa chỉ thường trú
    ws.column(23).setWidth(16);//Tài khoản chính
    ws.column(24).setWidth(15);//Hạn sử dụng
    ws.column(25).setWidth(65);//Thuê bao trả trước được tham gia khuyến mại
    ws.column(26).setWidth(65);//Gói cước trả trước ưu tiên mời KH đăng ký

    //dịch vụ 3G
    //dịch vụ 1
    ws.column(27).setWidth(17);//Mã DV1
    ws.column(28).setWidth(30);//Gói 3g 1
    ws.column(29).setWidth(25);//Ngày bắt đầu dịch vụ 1
    ws.column(30).setWidth(25);//Ngày kết thúc dịch vụ 1
    ws.column(31).setWidth(10);//Gia hạn 1

    //dịch vụ 2
    ws.column(32).setWidth(17);//Mã DV2
    ws.column(33).setWidth(30);//Gói 3g 2
    ws.column(34).setWidth(25);//Ngày bắt đầu dịch vụ 2
    ws.column(35).setWidth(25);//Ngày kết thúc dịch vụ 2
    ws.column(36).setWidth(10);//Gia hạn 2

    //dịch vụ 3
    ws.column(37).setWidth(17);//Mã DV3
    ws.column(38).setWidth(30);//Gói 3g 3
    ws.column(39).setWidth(25);//Ngày bắt đầu dịch vụ 3
    ws.column(40).setWidth(25);//Ngày kết thúc dịch vụ 3
    ws.column(41).setWidth(10);//Gia hạn 3

    xlStyleSmall = wb.createStyle({
        alignment: {
            vertical: ['center'],
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            name: 'Arial',
            color: '#324b73',
            size: 12,
        }
    });

    xlStyleBig = wb.createStyle({
        alignment: {
            vertical: ['center'],
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            name: 'Arial',
            color: '#324b73',
            size: 12,
        }
    });

    xlStyleNone = wb.createStyle({
        alignment: {
            vertical: ['center'],
            horizontal: ['center'],
            wrapText: true,
        },
        font: {
            bold: true,
            name: 'Arial',
            color: '#324b73',
            size: 12,
        },
    });

    fileNamexlxs = "(" + cTimee.getHours() + " Gio -" + cTimee.getMinutes() + " Phut Ngay " + cTimee.getDate() + " Thang " + (cTimee.getMonth() + 1) + " Nam " + cTimee.getFullYear() + ")   " + fileNametxt + ".xlsx";

    let header = [
        "STT",
        "Số thuê bao",
        "MSIN",
        "Loại thuê bao",
        "Gọi đi",
        "Gọi đến",
        "Loại SIM",
        "Hạng hội viên",
        "Tỉnh",
        "Ngày KH",
        "Mã KH",
        "Mã CQ",
        "Tên thuê bao",
        "Ngày sinh",
        "Số GT",
        "Ngày cấp",
        "Số PIN/PUK",
        "Số PIN2/PUK2",
        "Đối tượng",
        "Địa chỉ chứng từ",
        "Địa chỉ thanh toán",
        "Địa chỉ thường trú",
        "Tài khoản chính",
        "Hạn sử dụng",
        "Thuê bao trả trước được tham gia khuyến mại",
        "Gói cước trả trước ưu tiên mời KH đăng ký",
        //dịch vụ 3G
        //dịch vụ 1
        "Mã DV1",
        "Gói 3g 1",
        "Ngày bắt đầu dịch vụ 1",
        "Ngày kết thúc dịch vụ 1",
        "Gia hạn 1",

        //dịch vụ 2
        "Mã DV2",
        "Gói 3g 2",
        "Ngày bắt đầu dịch vụ 2",
        "Ngày kết thúc dịch vụ 2",
        "Gia hạn 2",

        //dịch vụ 3
        "Mã DV3",
        "Gói 3g 3",
        "Ngày bắt đầu dịch vụ 3",
        "Ngày kết thúc dịch vụ 3",
        "Gia hạn 3",
    ];

    for (let i = 0; i < header.length; i++) {
        await mainWindow.webContents.send(crawlCommand.log, "vòng for ghi header  " + i + " title " + headeTitle + "-" + header[i]);
        await writeToXcell(1, Number.parseInt(i) + 1, headeTitle + "-" + header[i]);
    }

    await mainWindow.webContents.send(crawlCommand.log, "puppeteeer file ouput tên là  " + fileNamexlxs);
    ws.row(1).setHeight(defaultHeight);
    startStartIndex = 0;
    await mainWindow.webContents.send(crawlCommand.hideBTN, true);
    callback();
}

function specialForOnlyHitButton() {
    prepareExxcel(doCrawl);
}

async function readFile(callback) {

    let arraySourceFileName = directionToSource.split("\\");
    let isNew = false;
    if (fileNametxt != "") {
        isNew = true;
    }
    //tách tên file
    fileNametxt = arraySourceFileName[arraySourceFileName.length - 1];
    let fileNametxtRemoveExxtension = fileNametxt.replace('.xlsx', '');
    // if (err) {
    //     ////console.log("An error ocurred reading the file :" + err.message);
    //     await mainWindow.webContents.send(crawlCommand.readError, fileNametxt.replace('.xlsx', ''));
    //     return;
    // }

    let workbook = XLSX.readFile(directionToSource);//
    let sheet_name_list = workbook.SheetNames; // laasy cacs sheet
    data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); //if you have multiple sheets

    if (data == '' || data == null) {
        await mainWindow.webContents.send(crawlCommand.readErrorNull, fileNametxtRemoveExxtension);
    } else {
        tResult = [];
        //assync 1 mảng
        await asyncReadFileExcel(data, function (item) {
            tResult.push(item);
        })
        console.log(tResult);
        await mainWindow.webContents.send(crawlCommand.log, 'dữ liệu trong tệp là  ' + tResult);

        if (isNew == true) {
            await mainWindow.webContents.send(crawlCommand.log, 'đọc tệp lần đầu tiên thành công tên tệp là ' + fileNametxtRemoveExxtension);
            await mainWindow.webContents.send(crawlCommand.readSuccess, fileNametxtRemoveExxtension);
        }
        else {
            await mainWindow.webContents.send(crawlCommand.log, 'đọc tệp lần nữa thành công tên tệp là ' + fileNametxtRemoveExxtension);
            await mainWindow.webContents.send(crawlCommand.readSuccessFirtTime, fileNametxtRemoveExxtension);
        }
        callback();
    }

    // fs.readFile(directionToSource, 'utf-8', async (err, data) => {

    //     // Change how to handle the file content
    //     if (data == '' || data == null) {
    //         await mainWindow.webContents.send(crawlCommand.readErrorNull, fileNametxt);
    //     } else {

    //     }
    // });
}

async function writeToFileXLSX() {
    await wb.write(fileNamexlxs);
}

async function writeToXcell(x, y, title) {
    //console.log("Ghi vao o ", x, y, "gia tri", title);
    await mainWindow.webContents.send(crawlCommand.log, "Ghi vao o " + x + ":" + y + " gia tri " + title);
    title += "";

    if (title.startsWith("header")) {
        let tTitle = title.split("-")[1];
        title = JSON.stringify(title);
        //title.replace("\"/g","");
        ws.cell(x, y).string(tTitle).style(xlStyleNone);
    } else {
        ws.cell(x, y).string(title).style(xlStyleSmall);
    }
    // }
}

async function writeNumberToCell(x, y, number) {
    await ws.cell(x, y).number(number).style(xlStyleSmall);
}

//sleep đi một vài giây
function timer(ms) {
    return new Promise(res => setTimeout(res, ms));
}

async function asyncReadFileExcel(array, callback) {
    for (let index = 0; index < array.length; index++) {
        await callback(array[index]["Số thuê bao"], index);
    }
}

async function asyncForEach(array, startIndex, callback) {
    let cIndex = 1;
    for (let index = startIndex; index < array.length; index++) {

        await mainWindow.webContents.send(crawlCommand.log, index + ' / ' + inputPhoneNumberArray.length);

        //đặt lại biến can write ặmc định là true
        canWrite = true;

        await callback(array[index], index);

        await mainWindow.webContents.send(crawlCommand.running, (index + 1) + " " + inputPhoneNumberArray.length);

        //sau ghi đủ số lượng threshHoldeCount sẽ ghi vào file excel
        if (index % threshHoldeCount === 0 && index > 0) {
            await mainWindow.webContents.send(crawlCommand.log, 'đax đạt đủ thresshold  ' + threshHoldeCount);
            await writeToFileXLSX();
        }

        //crawl xong 1 số -> nghỉ await timer(delayInMilliseconds);
        if (index < array.length - 1) {
            await mainWindow.webContents.send(crawlCommand.log, 'delay   ' + delayInMilliseconds);
            await timer(delayInMilliseconds);
        }
    }
}

//crawl

function doLogin(_username, _password) {
    concurentLogin = null;
    //đang login
    concurentLogin = puppeteer.launch({ headless: false, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
        mainBrowser = browser;
        pageLogin = await browser.newPage();

        await mainWindow.webContents.send(crawlCommand.loginSuccess, 2);
        await mainWindow.webContents.send(crawlCommand.log, 'doLogin');
        
        await pageLogin.goto(crawlUrl);//, { waitUntil: 'networkidle0' });

        pageLogin.setViewport({ width: 2600, height: 3000 });
        
        //có dialog hiệnh lên
        //hầu hết các lỗi dialog, -> đóng trình duệt
        //dialog số không hợp lệ(sai định dang số, số quá ngắn, quá dài hoặc otp bị sai), không đóng google
        pageLogin.on('dialog', async dialog => {

            let mssg = dialog.message();

            // await pageLogin.evaluate(({ mssg }) => {

            //     console.log('puppeteer alert wwith: ', mssg);

            // }, { mssg });

            console.log('puppeteer alert wwith: ', mssg);
            await mainWindow.webContents.send(crawlCommand.log, 'puppeteer alert with: ' + mssg);
            //await mainWindow.webContents.send('crawl:incorrect_number', inputPhoneNumberArray[cIII]);
            if (dialog.message() == wrongLogin) {
                await mainWindow.webContents.send(crawlCommand.log, 'wrongLogin');
                await mainWindow.webContents.send(crawlCommand.loginSuccess, 0);
                await browser.close();
                concurentLogin = null;
            } else if (dialog.message() == wrongPassword) {
                await mainWindow.webContents.send(crawlCommand.log, 'wrongPassword');
                await mainWindow.webContents.send(crawlCommand.loginSuccess, -2);
                await browser.close();
                concurentLogin = null;
            } else if (dialog.message() == lackPassword) {
                await mainWindow.webContents.send(crawlCommand.log, 'lackPassword');
                await mainWindow.webContents.send(crawlCommand.loginSuccess, -3);
                await browser.close();
                concurentLogin = null;
            } else if (dialog.message() == wrongOTP) {
                await mainWindow.webContents.send(crawlCommand.otp, 0);
                await mainWindow.webContents.send(crawlCommand.log, 'wrongOTP');
            } else if (dialog.message() == timoutOTP) {
                //phiên kiểm tra hết hạn, đóng trình duyệt mở lại login
                await mainWindow.webContents.send(crawlCommand.otp, -2);
                await mainWindow.webContents.send(crawlCommand.log, 'timoutOTP');
                await browser.close();
                concurentLogin = null;
            }
            //phần crawl dữ liệu, có dialog số điện thoại không hợp lệ(số dài hơn hoặc ngắn hơn quy định)
            else if (dialog.message() == wrongPhoneNumber) {
                await mainWindow.webContents.send(crawlCommand.signalWrite, -1);
                canWrite = false;
            }
            else { // dialog có nội dung chưa biết
                await mainWindow.webContents.send(crawlCommand.loginSuccess, -1);
                //await mainWindow.webContents.send(crawlCommand.otp, -1);
                await mainWindow.webContents.send(crawlCommand.log, 'alert unknown exception ' + mssg);
                await dialog.dismiss();
                await mainBrowser.close();
                concurentLogin = null;
            }
            dialog.dismiss();
        });

        await pageLogin.waitForNavigation({ waitUntil: 'networkidle0' });

        // await pageLogin.evaluate(({ _username, _password }) => {
        //     document.getElementById("txtUsername").value = _username;
        //     document.getElementById("txtPassword").value = _password;
        //     document.getElementById("btnLogin").click;
        // }, { _username, _password });

        await pageLogin.$eval('body #ctl01 .wrap-body .inner .tbl-login #txtUsername', (el, value) => el.value = value, _username);
        await pageLogin.$eval('body #ctl01 .wrap-body .inner .tbl-login #txtPassword', (el, value) => el.value = value, _password);


        //ngăn race condition
        await Promise.all([pageLogin.click('#ctl01 .wrap-login .inner .tbl-login #btnLogin'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);

        //await mainWindow.webContents.send(crawlCommand.log, 'waiting login');


        //đăng nhập thành công
        await mainWindow.webContents.send(crawlCommand.loginSuccess, 1);

        //mật khẩu OTP
        //otp
        ipcMain.on(crawlCommand.otp, async function (e, item) {


            //đang xác thực OTP
            await mainWindow.webContents.send(crawlCommand.otp, 2);

            await mainWindow.webContents.send(crawlCommand.log, 'otp ' + item);

            await pageLogin.$eval('#ctl01 .wrap-body .inner .tbl-login #txtOtp', (el, value) => el.value = value, item);

            //ngăn race condition
            await Promise.all([pageLogin.click('#ctl01 .wrap-body .inner .tbl-login #btnProcess'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);

            // await pageLogin.evaluate(({ ite }) => {
            //     document.getElementById("txtOtp").value = ite;
            //     document.getElementById("btnProcess").click;
            // }, { item });

            await mainWindow.webContents.send(crawlCommand.log, 'waiting otp');



            //xác thực mật khẩu otp thành công
            await mainWindow.webContents.send(crawlCommand.otp, 1);

            // await mainBrowser.close();
            // concurentLogin = null;
        });


        //crawl data
        ipcMain.on(crawlCommand.doCrawl, async function (e, item) {
            ////console.log(e, item);
            delayInMilliseconds = item == null ? 10000 : item;
            //console.log("delayInMilliseconds", delayInMilliseconds,"directionToSource",directionToSource);
            await mainWindow.webContents.send(crawlCommand.log, 'bấm crawl đường dẫn đến thư mục ' + directionToSource);
            if (directionToSource == "" || directionToSource == null) {
                await chooseSource(readFile, specialForOnlyHitButton);
            } else {
                prepareExxcel(doCrawl);
            }

        })


    }).catch(async (err, browser) => {
        //các trường hợp do user đóng app, hoặc do mất mạng
        await mainWindow.webContents.send(crawlCommand.loginSuccess, -1);
        //await mainWindow.webContents.send(crawlCommand.otp, -1);
        await mainWindow.webContents.send(crawlCommand.result, false);

        await mainWindow.webContents.send(crawlCommand.log, 'uncaught exception ' + err);
        await mainBrowser.close();
        concurentLogin = null;
    });
}

async function doCrawl() {

    //await page.goto(crawlUrl);
    await mainWindow.webContents.send(crawlCommand.log, 'bắt đầu crawl ');
    const start = async () => {
        await asyncForEach(inputPhoneNumberArray, startStartIndex, async (element, index) => {

            await mainWindow.webContents.send(crawlCommand.log, 'crawl đến phần tử thứ  ' + index + " là số thuê bao " + inputPhoneNumberArray[index] + " = " + element);

            //nhập vào số điện thoại
            //await pageLogin.$eval('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtThueBao', (el, value) => el.value = value, element);

            await pageLogin.evaluate(({ element }) => {
                document.getElementById("txtThueBao").value = element;
                document.getElementById("btnSearch").click();
            }, { element });

            //ngăn race condition
            //await Promise.all([pageLogin.click('#ctl01 .wrap-body .inner .tbl-login #btnProcess'), pageLogin.waitForNavigation({ waitUntil: 'networkidle0' })]);

            //bấm nút tìm
            //await pageLogin.click('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #btnSearch');

            //đợi page load
            //cần xử lý sleep vài giây
            await timer(1100);

            //làm mới mảng curent Data
            currentData.length = 0;

            //1
            currentData.push(index + 1);// sso thứ tự
            //2
            currentData.push(inputPhoneNumberArray[index]);//"Số thuê bao",
            //3
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtMSIN').getProperty('innerHTML')).jsonValue());//"MSIN",
            //4
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtLoaiTB').getProperty('innerHTML')).jsonValue());//"Loại thuê bao",
            //5
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #chkGoiDi').getProperty('innerHTML')).jsonValue() === true ? mCheckTrue : mCheckFalse);//"Gọi đi",
            //6
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #chkGoiDen').getProperty('innerHTML')).jsonValue() === true ? mCheckTrue : mCheckFalse);//"Gọi đến",
            //7
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtSimType').getProperty('innerHTML')).jsonValue());//"Loại SIM",
            //8
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtHangHoiVien').getProperty('innerHTML')).jsonValue());//"Hạng hội viên",
            //9
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtTinh').getProperty('innerHTML')).jsonValue());//"Tỉnh",
            //10
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtNgayKH').getProperty('innerHTML')).jsonValue());//"Ngày KH",
            //11
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtMaKH').getProperty('innerHTML')).jsonValue());//"Mã KH",
            //12
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtMaCQ').getProperty('innerHTML')).jsonValue());//"Mã CQ",
            //13
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtTB').getProperty('innerHTML')).jsonValue());//"Tên thuê bao",
            //14
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtNgaySinh').getProperty('innerHTML')).jsonValue());//"Ngày sinh",
            //15
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtSoGT').getProperty('innerHTML')).jsonValue());//"Số GT",
            //16
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtNoiCap').getProperty('innerHTML')).jsonValue());//"Ngày cấp",
            let pinpuk = await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtPIN') + "/"
                + await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtPUK');
            //17
            currentData.push(pinpuk);//"Số PIN/PUK",

            let pinpuk2 = await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtPIN2') + "/"
                + await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtPUK2');
            //18
            currentData.push(pinpuk2);//"Số PIN2/PUK2",
            //19
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtDoiTuong').getProperty('innerHTML')).jsonValue());//"Đối tượng",
            //20
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtDiaChiChungTu').getProperty('innerHTML')).jsonValue());//"Địa chỉ chứng từ",
            //21
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtDiaChiThanhToan').getProperty('innerHTML')).jsonValue());//"Địa chỉ thanh toán",
            //22
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtDiaChiThuongTru').getProperty('innerHTML')).jsonValue());//"Địa chỉ thường trú",
            //23
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtTKC').getProperty('innerHTML')).jsonValue());//"Tài khoản chính",
            //24
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtHSD').getProperty('innerHTML')).jsonValue());//"Hạn sử dụng",
            //25
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtKhuyenMai').getProperty('innerHTML')).jsonValue());//"Thuê bao trả trước được tham gia khuyến mại",
            //26
            currentData.push(await (await pageLogin.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtKhuyenNghi').getProperty('innerHTML')).jsonValue());//"Gói cước trả trước ưu tiên mời KH đăng ký",

            //bấm vào 3g tab
            await pageLogin.$x("//span[contains(., 'Lịch sử 3G')]");

            //sleep đi 1 giây
            await timer(1100);

            let dataFromTable = await pageLogin.$$eval('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor .box5 #tabContent table tr td', tableData => tableData.map((td) => {
                return td.innerHTML;
            }));

            await mainWindow.webContents.send(crawlCommand.log, "đọc từ dịch vụ 3g " + dataFromTable);

            if (canWrite) {
                //phần ghi ra file excel
                //đến phẩn tử 26 là hết phần thông tin khách
                let outerIndex = index;
                for (let index = 0; index < 25; index++) {
                    await writeToXcell(outerIndex + rowSpacing, index + 1, currentData[index]);
                }


                let currentCollumn = 26;
                for (let index = 0; index < dataFromTable.length; index++) {
                    //dataFromTable
                    if (index % breakPerSerrvice == 0) {
                        continue;
                    } else {
                        await writeToXcell(index + rowSpacing, currentCollumn + index, dataFromTable);
                    }
                }
            }
        });



        console.log("end");
        //lần chạy cuối cùng
        await writeToFileXLSX();

        await browser.close();

        await mainWindow.webContents.send(crawlCommand.result, true);

        concurentPup = null;
        //crawling = false;
    };

    start();
}

//liên lạc giữa index.js và index html
//open file
ipcMain.on(crawlCommand.openFile, async function (e, item) {
    chooseSource(readFile, nothing);
});

//login
ipcMain.on(crawlCommand.login, async function (e, item) {
    username = item.split(" ")[0];
    password = item.split(" ")[1];
    doLogin(username, password);
});