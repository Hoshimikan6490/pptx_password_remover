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
  const tagIndex = xml.indexOf('<p:modifyVerifier');
  if (tagIndex === -1) {
    return null;
  }

  const tagEndIndex = xml.indexOf('>', tagIndex);
  if (tagEndIndex === -1) {
    return null;
  }

  return `${xml.slice(0, tagIndex)}${xml.slice(tagEndIndex + 1)}`;
}

// [Content_Types].xml内のPPSX用ContentTypeをPPTX用に変換する関数
function convert_slideshow_to_presentation(contentTypesXml) {
  return contentTypesXml.replace(
    /application\/vnd\.openxmlformats-officedocument\.presentationml\.slideshow\.main\+xml/g,
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  );
}

// POSTリクエストを処理するルート
app.post('/api/progress', async function (req, res) {
  const isRemovePassword = req.query.type === 'removePassword';
  const isConvertToPPTX = req.query.type === 'convertToPPTX';

  // どちらかのクエリパラメータは必須
  if (!isRemovePassword && !isConvertToPPTX) {
    res
      .status(400)
      .send(
        'Specify action with query parameter: ?removePassword or ?convertToPPTX',
      );
    return;
  }

  // 両方同時にリクエストされた場合はエラー
  if (isRemovePassword && isConvertToPPTX) {
    res.status(400).send('Specify only one action at a time.');
    return;
  }

  // ファイルがアップロードされているか確認
  if (!req.files || Object.keys(req.files).length === 0) {
    res.status(400).send('No files were uploaded.');
    return;
  }

  // アップロードされたファイルがPowerPointファイルか確認
  let file = req.files.file;
  const extension = path.extname(file.name).toLowerCase();
  if (extension !== '.pptx' && extension !== '.ppsx') {
    res.status(400).send('Only .pptx or .ppsx files are supported.');
    return;
  }

  // 変換の場合は、.ppsxファイルであることを確認
  if (isConvertToPPTX && extension !== '.ppsx') {
    res.status(400).send('convertToPPTX mode only supports .ppsx files.');
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

  // パスワード削除の場合
  if (isRemovePassword) {
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
    if (newXmlString === null) {
      res
        .status(400)
        .send(
          'Uploaded PowerPoint file does not appear to be password-protected.',
        );
      return;
    }
    zip.file('ppt/presentation.xml', newXmlString);
  }

  // PPSXをPPTXに変換する場合
  if (isConvertToPPTX) {
    // [Content_Types].xmlを取得
    const contentTypesEntry = zip.file('[Content_Types].xml');
    if (!contentTypesEntry) {
      res
        .status(400)
        .send('Invalid PowerPoint file: [Content_Types].xml not found.');
      return;
    }

    // PPSX用ContentTypeをPPTX用に変換して、[Content_Types].xmlを更新
    const contentTypesXml = await contentTypesEntry.async('string');
    const convertedXml = convert_slideshow_to_presentation(contentTypesXml);
    zip.file('[Content_Types].xml', convertedXml);
  }

  // 修正されたPowerPointファイルを再構築して保存
  const rebuiltBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  await fsp.writeFile(uploadPath, rebuiltBuffer);

  // クライアントに修正されたPowerPointファイルをダウンロードさせる
  const outputExtension = isConvertToPPTX ? '.pptx' : extension;
  const suffix = isConvertToPPTX ? '_convertedToPPTX' : '_passwordRemoved';
  const fileName = `${path.basename(file.name, extension)}${suffix}${outputExtension}`;
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
