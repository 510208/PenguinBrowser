# PenguinBrowser

[![GitHub stars](https://img.shields.io/github/stars/510208/PenguinBrowser?color=brightgreen&style=for-the-badge)](https://github.com/510208/NotUseComputer/)
[![](https://img.shields.io/badge/Blog-510208's%20Blog-brightgreen?style=for-the-badge&logo=appveyor)](https://sam0616.pixnet.net)
![MIT license](https://img.shields.io/badge/license-MIT-brightgreen.svg?style=for-the-badge&logo=appveyor)
[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/X8X4CZE3V)

## 🔲 目錄

[💯 關於](#-關於)

[❌ 無法使用]

[😍 關於我](#-關於我)

[📄 授權](#-授權)

## 💯 關於

此軟體為簡易版瀏覽器，透過IE11(或12)內核建立網頁連線，倘若您的系統並未啟用IE，建議先行啟用或自行編譯適用版本，否則可能造成主程式無法運行。

## ❌ 無法使用時的除錯方式

### 註冊表存取錯誤
如果出現如下問題，請照以下教學處理：

![https://img.onl/LTtbKW](https://img.onl/LTtbKW)

將以下代碼複製成reg檔案，並且運行：

```
Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\TypeLib\{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}]

[HKEY_CLASSES_ROOT\TypeLib\{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}\1.1]
@="Microsoft Internet Controls"

[HKEY_CLASSES_ROOT\TypeLib\{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}\1.1\0]

[HKEY_CLASSES_ROOT\TypeLib\{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}\1.1\0\win32]
@="C:\\WINDOWS\\system32\\ieframe.dll"
```

或是下載連結：
[💬 下載註冊檔](/Error-Debug/Browser_Error.reg)

### IE未啟用

倘若您的IE11未啟用，請依下圖操作：

![Step1](https://img.onl/2n2dUB)

進入控制台>程式和功能

![Step2](https://img.onl/NxLEGN)

開啟或關閉Windows功能

![Step3](https://img.onl/NkXD9l)

- [x] Internet Explorer 11

![Step4](https://img.onl/NkXD9l)

確定咩~

___

或是直接安裝IE11：
[💬 IE11連結](/IE11/IE11_Link.md)

## 😍 關於我

[![YouTube](https://img.shields.io/badge/YouTube-%E8%B7%9F%E8%91%97%E4%BC%81%E9%B5%9D%E5%93%A5%E5%AD%B8%E9%9B%BB%E8%85%A6-red?style=for-the-badge&logo=appveyor)](https://www.youtube.com/channel/UC6orwHdQNVzwHsA6M7HYD9g/videos?view=0&sort=p&shelf_id=0)
[![Blog](https://img.shields.io/badge/Pixnet-%E8%B7%9F%E8%91%97%E4%BC%81%E9%B5%9D%E5%93%A5%E5%AD%B8%E9%9B%BB%E8%85%A6-blue?style=for-the-badge)](https://sam0616.pixnet.net)
[![WordPress](https://img.shields.io/badge/WordPress-%E8%B7%9F%E8%91%97%E4%BC%81%E9%B5%9D%E5%93%A5%E5%AD%B8%E9%9B%BB%E8%85%A6-yellowgreen?style=for-the-badge&logo=appveyor)](https://510208.nde.tw)

喜歡這支程式，可以訂閱點讚開小鈴鐺、部落格新增到書籤等...^^

## 📄 授權

我忘記設定了，反正此軟體的授權我是用GNU License v3.0，詳見軟體內或
[GNU License v3.0](https://www.gnu.org/licenses/gpl-3.0.zh-tw.html)

[中文版GNU License v3.0](/LICENSE_ZH.md)
