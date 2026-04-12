const express = require('express');
const path = require('path');
const fileUpload = require('express-fileupload');
const fs = require('fs');
const fsp = fs.promises;
const JSZip = require('jszip');

const app = express();
const port = 6490;

// デバッグモードの切り替え
// trueの場合、アップロードしたパワポのファイルを削除しないようにする。
const debugMode = false;

// ファイルアップロードの設定
app.use(
  fileUpload({
    defCharset: 'utf8',
    defParamCharset: 'utf8',
  }),
);

// 静的ファイルの提供
app.use(express.static(path.join(__dirname, 'public')));

// ルートパスへのGETリクエストに対して、index.htmlを返す
app.get('/', (req, res) =>
  res.sendFile(path.join(__dirname, 'public/pages/index.html')),
);

// presentation.xmlからp:modifyVerifierを削除する関数
function remove_p_modifyVerifier(xml) {
  let i = 0;
  while (1) {
    i++;
    if (xml.substr(i, 1) == 'p') {
      if (xml.substr(i + 2, 14) == 'modifyVerifier') {
        var first_hand = xml.substr(0, i - 1);

        for (j = i; i <= 6000; j++) {
          if (xml.substr(j, 1) == '>') {
            var latter_hand = xml.slice(j + 1);
            break;
          }
        }

        return first_hand + latter_hand;
      }
    }
  }
}

// POSTリクエストを処理するルート
app.post('/', async function (req, res) {
  // ファイルがアップロードされているか確認
  if (!req.files || Object.keys(req.files).length === 0) {
    res.status(400).send('No files were uploaded.');
    return;
  }

  // アップロードされたファイルがPowerPointファイルか確認
  let file = req.files.file;
  const extension = path.extname(file.name).toLowerCase();
  if (extension !== '.pptx' && extension !== '.ppsx') {
    res.status(400).send('Only .pptx and .ppsx files are supported.');
    return;
  }

  // ファイルをサーバーに保存
  const uploadPath = path.join(__dirname, 'public/uploads', file.name);
  await file.mv(uploadPath);

  // アップロードされたPowerPointのファイルを読み込み
  const originalBuffer = await fsp.readFile(uploadPath);
  const zip = await JSZip.loadAsync(originalBuffer, {
    checkCRC32: false,
  });

  // presentation.xmlを取得
  const presentationEntry = zip.file('ppt/presentation.xml');
  if (!presentationEntry) {
    res
      .status(400)
      .send('Invalid PowerPoint file: presentation.xml not found.');
    return;
  }

  // p:modifyVerifierを削除して、presentation.xmlを更新
  const xmlString = await presentationEntry.async('string');
  const newXmlString = remove_p_modifyVerifier(xmlString);
  zip.file('ppt/presentation.xml', newXmlString);

  // 修正されたPowerPointファイルを再構築して保存
  const rebuiltBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  await fsp.writeFile(uploadPath, rebuiltBuffer);

  // クライアントに修正されたPowerPointファイルをダウンロードさせる
  const fileName = `${path.basename(file.name, extension)}_passwordRemoved${extension}`;
  res.download(
    path.join(__dirname, `public/uploads/${file.name}`),
    fileName,
    function (err) {
      if (err) {
        console.log(err);
      } else {
        if (!debugMode) fs.unlinkSync(`public/uploads/${file.name}`);
      }
    },
  );
});

app.listen(port, function () {
  console.log(`起動しました！　http://localhost:${port} をご覧ください!`);
});
