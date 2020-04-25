const electron = require('electron');
const puppeteer = require('puppeteer');
var fs = require('fs');
var xl = require('excel4node');

const { app, BrowserWindow, ipcMain, Menu, dialog } = electron;

var concurentPup;
var delayInMilliseconds = 10000;
var inputPhoneNumberArray = [];
let fileNametxt = "";
var wb;
var ws;

var wrongPassword = "Mật khẩu phải ít nhất 8 ký tự, 1 ký tự hoa, 1 ký tự đặc biệt, 1 ký tự số. Nếu không hợp lệ vui lòng đổi trước khi đăng nhập";
var wrongLogin = "Tài khoản không hợp lệ, vui lòng thử lại";
var wrongOTP = "Mã OTP không hợp lệ, vui lòng thử lại";
var wrongPhoneNumber = "Số không hợp lệ";

var crawlUrl = "http://10.149.34.250:1609/Views/KhachHang/ThongTinKhachHang.aspx"; // vì nếu chưa dăng nhập thì vào trang lấy thông tin khách hàng cũng sẽ bị redirect về trang đăng nhập

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800, height: 600, webPreferences: {
            nodeIntegration: true
        }
    });

    mainWindow.loadURL(`file://${__dirname}/index.html`);

    // Build menu from template
    const mainMenu = Menu.buildFromTemplate(mainMenuTemplate);
    // Insert menu
    Menu.setApplicationMenu(mainMenu);

    mainWindow.on('closed', function () {
        mainWindow = null;
    })
}

app.on('ready', createWindow);


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