const electron = require('electron');
var cheerio = require('cheerio');
var superagent = require('superagent');
var json2xls = require('json2xls');
var fs = require('fs');
var async = require('async');

// 控制应用生命周期的模块
const {app} = electron;
// 创建本地浏览器窗口的模块
const {BrowserWindow} = electron;

// 指向窗口对象的一个全局引用，如果没有这个引用，那么当该javascript对象被垃圾回收的 时候该窗口将会自动关闭
let win;

const cookie = 'you cookie';
const urlTemp = 'you url';
const urlCash = 'you url';

var urlQueryStrings = [],
  items = [],
  urlStringsCash = [],
  itemsCash = [];

function createWindow() {
  // 创建一个新的浏览器窗口
  win = new BrowserWindow({width: 800, height: 600});

  // 并且装载应用的index.html页面
  win.loadURL(`file://${__dirname}/index.html`);

  // 打开开发工具页面 
  //win.webContents.openDevTools(); 
  
  //隐藏菜单栏
  // win.setMenuBarVisibility(false); 
  
  //当窗口关闭时调用的方法
  win.on('closed', () => {
    // 解除窗口对象的引用，通常而言如果应用支持多个窗口的话，你会在一个数组里 存放窗口对象，在窗口关闭的时候应当删除相应的元素。
    win = null;
  });
}

// 当Electron完成初始化并且已经创建了浏览器窗口，则该方法将会被调用。 有些API只能在该事件发生后才能被使用。
app.on('ready', createWindow);

// 当所有的窗口被关闭后退出应用
app.on('window-all-closed', () => {
  // 对于OS X系统，应用和相应的菜单栏会一直激活直到用户通过Cmd + Q显式退出
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  // 对于OS X系统，当dock图标被点击后会重新创建一个app窗口，并且不会有其他 窗口打开
  if (win === null) {
    createWindow();
  }
});

// 在这个文件后面你可以直接包含你应用特定的由主进程运行的代码。 也可以把这些代码放在另一个文件中然后在这里导入。 和子进程之间异步通信
const ipcMain = require('electron').ipcMain;

ipcMain.on('async-message', function (event, arg) {
  //订单管理
  if (arg === 'orderSearch') {
    superagent
      .get(urlTemp)
      .set('Cookie', cookie)
      .end(function (err, sres) {

        //将页面html传给cheerio解析
        var $ = cheerio.load(sres.text);
        //获取当前菜单下明细资料的总页数
        var pageCount = Number($('.page').children().first().text().split('/')[1]);
        //根据总页数构造url查询参数
        for (var i = 1; i < pageCount + 1; i++) {
          urlQueryStrings.push(urlTemp + '?xm=&page=' + i);
        }
        //异步并发处理其他的url
        asyncFun(event);
      }
      );
  } else if (arg === 'cashSearch') {//提现记录
    superagent
      .get(urlCash)
      .set('Cookie', cookie)
      .end(function (err, sres) {

        //将页面html传给cheerio解析
        var $ = cheerio.load(sres.text);
        //获取当前菜单下明细资料的总页数
        var pageCount = Number($('.page').children().first().text().split('/')[1]);
        //根据总页数构造url查询参数
        for (var i = 1; i < pageCount + 1; i++) {
          urlStringsCash.push(urlCash + '?page=' + i);
        }
        //异步并发处理其他的url
        asyncFunCash(event);
      });
  }
});

//对url进行处理
var fetchUrl = function (url, callback) {
  superagent
    .get(url)
    .set('Cookie', cookie)
    .end(function (err, sres) {

      callback(null, sres.text);
    })

};

//并发数控制在3以内，来执行异步请求
function asyncFun(event) {
  async
    .mapLimit(urlQueryStrings, 3, function (url, callback) {
      fetchUrl(url, callback);
    }, function (err, result) {
      for (let i = 0; i < result.length; i++) {
        var $ = cheerio.load(result[i]);
        $('tr[bgcolor=#cccccc]').each(function (i, elem) {
          items.push({
            '会员姓名': $(this)
              .children('td')
              .eq(4)
              .text(),
            '交易金额': $(this)
              .children('td')
              .eq(5)
              .text(),
            '数量': $(this)
              .children('td')
              .eq(6)
              .text(),
            '订单日期': $(this)
              .children('td')
              .eq(8)
              .text(),
            '商家用户': $(this)
              .children('td')
              .eq(9)
              .text()
          });
        });
      }

      //将内容生成Excel
      var xls = json2xls(items);
      fs.writeFileSync('订单管理.xlsx', xls, 'binary');

      //发送complete到子线程，通知操作完成
      event
        .sender
        .send('async-reply', 'complete');

    })
}

//cash
function asyncFunCash(event) {
  async
    .mapLimit(urlStringsCash, 3, function (url, callback) {
      fetchUrl(url, callback);
    }, function (err, result) {
      for (let i = 0; i < result.length; i++) {
        var $ = cheerio.load(result[i]);
        $('tr[bgcolor=#cccccc]').each(function (i, elem) {
          itemsCash.push({
            'ID': $(this)
              .children('td')
              .eq(0)
              .text(),
            '提现日期': $(this)
              .children('td')
              .eq(1)
              .text(),
            '用户名': $(this)
              .children('td')
              .eq(2)
              .text(),
            '金额': $(this)
              .children('td')
              .eq(3)
              .text(),
            '姓名': $(this)
              .children('td')
              .eq(8)
              .children('span')[0]
              .firstChild
              .data
          });
        });
      }

      //将内容生成Excel
      var xls = json2xls(itemsCash);
      fs.writeFileSync('提现记录.xlsx', xls, 'binary');

      //发送complete到子线程，通知操作完成
      event
        .sender
        .send('async-reply', 'complete');

    })
}