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
var rowSpacing = 2;
var directionToSource = "";
var wrongPassword = "Mật khẩu phải ít nhất 8 ký tự, 1 ký tự hoa, 1 ký tự đặc biệt, 1 ký tự số. Nếu không hợp lệ vui lòng đổi trước khi đăng nhập";
var lackPassword = "Mật khẩu";
var wrongLogin = "Tài khoản không hợp lệ, vui lòng thử lại";

var wrongOTP = "Mã OTP không hợp lệ, vui lòng thử lại";
var wrongPhoneNumber = "Số không hợp lệ";
var timoutOTP = "Phiên kiểm tra Otp hết hạn, hệ thống trở về trang đăng nhập.... ";

var crawlUrl = "http://10.149.34.250:1609/Views/KhachHang/ThongTinKhachHang.aspx"; // vì nếu chưa dăng nhập thì vào trang lấy thông tin khách hàng cũng sẽ bị redirect về trang đăng nhập

var crawlCommand = {
    login: "crawl:login",
    otp: "crawl:otp",
    openFile: "crawl:openFile",
    wrongPhoneNumber: "crawl:incorrect_number",
    hideBTN: "crawl:hideBTN",
    networkError: "crawl:network_error",
    running: "crawl:onrunning",
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
    loginSuccess: "crawl:login_success"
};

var page, pageLogin;
var breakPerSerrvice = 6;//có 6 cột dịch vụ
function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800, height: 600, webPreferences: {
            nodeIntegration: true
        }
    });

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

    let header = ["STT",
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
        writeToXcell(1, Number.parseInt(i) + 1, headeTitle + "-" + header[i]);
    }

    ws.row(1).setHeight(defaultHeight);
    startStartIndex = 0;
    mainWindow.webContents.send('crawl:hideBTN', true);
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

    let workbook = XLSX.readFile(fileNametxt);//
    let sheet_name_list = workbook.SheetNames; // laasy cacs sheet
    data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); //if you have multiple sheets

    if (data == '' || data == null) {
        await mainWindow.webContents.send(crawlCommand.readErrorNull, fileNametxtRemoveExxtension);
    } else {
        tResult = [];
        //assync 1 mangr
        await asyncReadFileExcel(data, function (item) {
            tResult.push(item);
        })

        console.log(tResult);

        if (isNew == true) {
            await mainWindow.webContents.send(crawlCommand.readSuccess, fileNametxtRemoveExxtension);
        }
        else {
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

function writeToXcell(x, y, title) {
    //console.log("Ghi vao o ", x, y, "gia tri", title);

    title += "";

    if (title.startsWith("header")) {
        let ttitle = title.split("-")[1];
        title = JSON.stringify(title);
        //title.replace("\"/g","");
        ws.cell(x, y).string(ttitle).style(xlStyleNone);
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
        await callback(array[index]['Số điện thoại'], index);
    }
}

async function asyncForEach(array, startIndex, callback) {
    let cIndex = 1;
    for (let index = startIndex; index < array.length; index++) {

        await callback(array[index], index);

        await mainWindow.webContents.send(crawlCommand.running, (index + 1) + " " + inputPhoneNumberArray.length);

        //sau ghi đủ số lượng threshHoldeCount sẽ ghi vào file excel
        if (index % threshHoldeCount === 0 && index > 0) {
            await writeToFileXLSX();
        }

        //crawl xong 1 số -> nghỉ await timer(delayInMilliseconds);
        if (index < array.length - 1) {
            await timer(delayInMilliseconds);
        }
    }
}

//crawl

function doLogin(_username, _password) {
    concurentLogin = null;
    concurentLogin = puppeteer.launch({ headless: false, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
        mainBrowser = browser;
        pageLogin = await mainBrowser.newPage();
        //đang login
        await mainWindow.webContents.send(crawlCommand.loginSuccess, 2);
        await pageLogin.goto(crawlUrl);
        pageLogin.setViewport({ width: 2600, height: 3000 });
        //có dialog hiệnh lên
        pageLogin.on('dialog', async dialog => {
            //await mainWindow.webContents.send('crawl:incorrect_number', inputPhoneNumberArray[cIII]);
            if (dialog.message() == wrongLogin) {
                await mainWindow.webContents.send(crawlCommand.loginSuccess, 0);
            } else if (dialog.message() == wrongPassword) {
                await mainWindow.webContents.send(crawlCommand.loginSuccess, -2);
            } else if (dialog.message() == lackPassword) {
                await mainWindow.webContents.send(crawlCommand.loginSuccess, -3);
            } else if (dialog.message() == wrongOTP) {
                await mainWindow.webContents.send(crawlCommand.otp, 0);
            } else if (dialog.message() == timoutOTP) {
                //phiên kiểm tra hết hạn, đóng trình duyệt mở lại login
                await mainWindow.webContents.send(crawlCommand.otp, -2);
                dialog.dismiss();
                await mainBrowser.close();
                concurentLogin = null;

            } else { // dialog có nội dung chưa biết
                await mainWindow.webContents.send(crawlCommand.loginSuccess, -1);
                await mainWindow.webContents.send(crawlCommand.otp, -1);
                console.log("uncaught exception");
                await mainBrowser.close();
                concurentLogin = null;
            }
            dialog.dismiss();
            await mainBrowser.close();
            concurentLogin = null;
        });

        await pageLogin.$eval('.wrap-body .inner .tbl-login #txtPassword', (el, value) => el.value = value, _username);
        await pageLogin.$eval('.wrap-body .inner .tbl-login #txtUsername', (el, value) => el.value = value, _password);

        await pageLogin.click('.wrap-login .inner .tbl-login #btnLogin');

        try {
            await pageLogin.waitForNavigation({ waitUntil: 'networkidle0' });
        } catch (e) {

        }

        await mainWindow.webContents.send(crawlCommand.loginSuccess, 1);

        //mật khẩu OTP
        //otp
        ipcMain.on(crawlCommand.otp, async function (e, item) {
            await pageLogin.$eval('#ctl01 .wrap-body .inner .tbl-login #txtOtp', (el, value) => el.value = value, item);

            await pageLogin.click('#ctl01 .wrap-body .inner .tbl-login #btnProcess');

            //đang xác thực OTP
            await mainWindow.webContents.send(crawlCommand.loginSuccess, 2);

            await mainBrowser.close();
            concurentLogin = null;
        });


    }).catch(async (err, browser) => {
        console.log("login error", err);
        await mainWindow.webContents.send(crawlCommand.loginSuccess, -1);
        await mainBrowser.close();
        concurentLogin = null;
    });
}

function doCrawl() {
    console.log("concurentPup", concurentPup != null);
    if (concurentPup != null) {
        //nếu đang mở, chắc là có lỗi từ phiên trước, đóng lại cho đữo tốn dung lượng
        concurentPup = null;
        // browser.close();
        page.close();
        doCrawl();
    } else {
        concurentPup = puppeteer.launch({ headless: true, executablePath: exPath == "" ? "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" : exPath }).then(async browser => {
            page = await browser.newPage();

            await page.goto(crawlUrl);
            page.setViewport({ width: 2600, height: 3000 });
            page.on('dialog', async dialog => {
                if (dialog.message() == wrongPhoneNumber) {
                    await mainWindow.webContents.send(crawlCommand.wrongPhoneNumber, inputPhoneNumberArray[cIII]);
                    //await dialog.dismiss();
                    await browser.close();
                    concurentPup = null;
                    dialog.dismiss();
                    startStartIndex = cIII + 1;

                    await writeToXcell(cIII + rowSpacing, 1, startStartIndex + "");
                    writeToXcell(cIII + rowSpacing, 3, inputPhoneNumberArray[cIII]);
                    writeToXcell(cIII + rowSpacing, 4, "");
                    writeToXcell(cIII + rowSpacing, 5, "");
                    writeToXcell(cIII + rowSpacing, 6, "" + 0);

                    await writeToXcell(cIII + rowSpacing, 8, wrongNumber);
                    ws.row(cIII + rowSpacing).setHeight(defaultHeight);
                    await doCrawl();
                }
            });


            //login
            await page.click('#btn-alert1 .effect-sadie');

            await page.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #username1', (el, value) => el.value = value, username);
            await page.$eval('#popupAlert1 #report .clearfix #form-login .from-login .form-row #password1', (el, value) => el.value = value, password);

            await page.click('#popupAlert1 #report .clearfix #form-login .from-login .form-row .button');

            await page.waitForNavigation({ waitUntil: 'networkidle0' })


            await page.goto(crawlUrl);

            const start = async () => {
                await asyncForEach(inputPhoneNumberArray, startStartIndex, async (element, index) => {

                    //nhập vào số điện thoại
                    await page.$eval('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtThueBao', (el, value) => el.value = value, inputPhoneNumberArray[index]);

                    //bấm nút tìm
                    await page.click('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #btnSearch');

                    //đợi page load
                    //cần xử lý sleep vài giây
                    await timer(1000);

                    await writeToXcell(index + rowSpacing, 1, index + 1);// sso thứ tự

                    await writeToXcell(index + rowSpacing, 2, inputPhoneNumberArray[index]);//"Số thuê bao",

                    await writeToXcell(index + rowSpacing, 3, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtMSIN').getProperty('innerHTML')).jsonValue());//"MSIN",
                    await writeToXcell(index + rowSpacing, 4, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtLoaiTB').getProperty('innerHTML')).jsonValue());//"Loại thuê bao",
                    await writeToXcell(index + rowSpacing, 5, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #chkGoiDi').getProperty('innerHTML')).jsonValue());//"Gọi đi",
                    await writeToXcell(index + rowSpacing, 6, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #chkGoiDen').getProperty('innerHTML')).jsonValue());//"Gọi đến",
                    await writeToXcell(index + rowSpacing, 7, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtSimType').getProperty('innerHTML')).jsonValue());//"Loại SIM",
                    await writeToXcell(index + rowSpacing, 8, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtHangHoiVien').getProperty('innerHTML')).jsonValue());//"Hạng hội viên",
                    await writeToXcell(index + rowSpacing, 9, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtTinh').getProperty('innerHTML')).jsonValue());//"Tỉnh",
                    await writeToXcell(index + rowSpacing, 10, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtNgayKH').getProperty('innerHTML')).jsonValue());//"Ngày KH",
                    await writeToXcell(index + rowSpacing, 11, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtMaKH').getProperty('innerHTML')).jsonValue());//"Mã KH",
                    await writeToXcell(index + rowSpacing, 12, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtMaCQ').getProperty('innerHTML')).jsonValue());//"Mã CQ",
                    await writeToXcell(index + rowSpacing, 13, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtTB').getProperty('innerHTML')).jsonValue());//"Tên thuê bao",
                    await writeToXcell(index + rowSpacing, 14, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtNgaySinh').getProperty('innerHTML')).jsonValue());//"Ngày sinh",
                    await writeToXcell(index + rowSpacing, 15, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtSoGT').getProperty('innerHTML')).jsonValue());//"Số GT",
                    await writeToXcell(index + rowSpacing, 16, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtNoiCap').getProperty('innerHTML')).jsonValue());//"Ngày cấp",
                    let pinpuk = await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtPIN')
                        + await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtPUK');
                    await writeToXcell(index + rowSpacing, 17, pinpuk);//"Số PIN/PUK",

                    let pinpuk2 = await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtPIN2')
                        + await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtPUK2');
                    await writeToXcell(index + rowSpacing, 18, pinpuk2);//"Số PIN2/PUK2",
                    await writeToXcell(index + rowSpacing, 19, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtDoiTuong').getProperty('innerHTML')).jsonValue());//"Đối tượng",
                    await writeToXcell(index + rowSpacing, 20, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtDiaChiChungTu').getProperty('innerHTML')).jsonValue());//"Địa chỉ chứng từ",
                    await writeToXcell(index + rowSpacing, 21, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtDiaChiThanhToan').getProperty('innerHTML')).jsonValue());//"Địa chỉ thanh toán",
                    await writeToXcell(index + rowSpacing, 22, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtDiaChiThuongTru').getProperty('innerHTML')).jsonValue());//"Địa chỉ thường trú",
                    await writeToXcell(index + rowSpacing, 23, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtTKC').getProperty('innerHTML')).jsonValue());//"Tài khoản chính",
                    await writeToXcell(index + rowSpacing, 24, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtHSD').getProperty('innerHTML')).jsonValue());//"Hạn sử dụng",
                    await writeToXcell(index + rowSpacing, 25, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtKhuyenMai').getProperty('innerHTML')).jsonValue());//"Thuê bao trả trước được tham gia khuyến mại",
                    await writeToXcell(index + rowSpacing, 26, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus #txtKhuyenNghi').getProperty('innerHTML')).jsonValue());//"Gói cước trả trước ưu tiên mời KH đăng ký",

                    //bấm vào 3g tab
                    await page.$x("//span[contains(., 'Lịch sử 3G')]");

                    //sleep đi 1 giây
                    await timer(1000);

                    let dataFromTable = await page.$$eval('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor .box5 #tabContent table tr td', tableData => tableData.map((td) => {
                        return td.innerHTML;
                    }));

                    let currentCollumn = 26;
                    for (let index = 0; index < dataFromTable.length; index++) {
                        //dataFromTable
                        if (index % breakPerSerrvice == 0) {
                            continue;
                        } else {
                            await writeToXcell(index + rowSpacing, currentCollumn + index, dataFromTable);
                        }
                    }

                    //cần nghĩ thêm
                    //dịch vụ 3G
                    //dịch vụ 1
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Mã DV1",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Gói 3g 1",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Ngày bắt đầu dịch vụ 1",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Ngày kết thúc dịch vụ 1",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Gia hạn 1",

                    // //dịch vụ 2
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Mã DV2",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Gói 3g 2",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Ngày bắt đầu dịch vụ 2",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Ngày kết thúc dịch vụ 2",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Gia hạn 2",

                    // //dịch vụ 3
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());// "Mã DV3",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Gói 3g 3",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Ngày bắt đầu dịch vụ 3",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Ngày kết thúc dịch vụ 3",
                    // await writeToXcell(index + rowSpacing, 1, await (await page.$$('#wraper #bodypage #col-right .tabs-wrap #rightarea #tracuuthongtinkhachhang .body .nobor .boxOB .midbox .nobor tbl-cus ').getProperty('innerHTML')).jsonValue());//"Gia hạn 3",

                });


                console.log("end");
                //lần chạy cuối cùng
                await writeToFileXLSX();

                await browser.close();

                await mainWindow.webContents.send('crawl:result', true);

                concurentPup = null;
                //crawling = false;
            }

            start();


        }).catch(async (err) => {
            //console.log("pupperteer error ", err);
            await mainWindow.webContents.send('crawl:network_error', true);
        });
    }
}

//liên lạc giữa index.js và index html
//crawl data
ipcMain.on(crawlCommand.doCrawl, async function (e, item) {
    ////console.log(e, item);
    delayInMilliseconds = item == null ? 10000 : item;
    //console.log("delayInMilliseconds", delayInMilliseconds,"directionToSource",directionToSource);
    if (directionToSource == "" || directionToSource == null) {
        await chooseSource(readFile, specialForOnlyHitButton);
    } else {
        prepareExxcel(doCrawl);
    }

})

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

