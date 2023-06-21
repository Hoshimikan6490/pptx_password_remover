# 実行する際に使用するコマンド
`npx nodemon index.js`

# xmlを編集する方法
`xml2js`パッケージを使う？？
```js
var parseString = require('xml2js').parseString;
var xml2js = require('xml2js');

var xml = `あなたのXML文字列`;

parseString(xml, function (err, result) {
    delete result.あなたのタグ名;
    var builder = new xml2js.Builder();
    var xml = builder.buildObject(result);
    console.log(xml);
});
```