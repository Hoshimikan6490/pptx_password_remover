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
      .status(500)
      .send('サーバーエラーが発生しました。サーバーのログを確認してください。');
    console.error(
      '処理種別が指定されていません。クエリパラメータで type=removePassword または type=convertToPPTX を指定してください。',
    );
    return;
  }

  // 両方同時にリクエストされた場合はエラー
  if (isRemovePassword && isConvertToPPTX) {
    res.status(500).send('サーバーエラーが発生しました。サーバーのログを確認してください。');
    console.error('処理種別は同時に複数指定できません。');
    return;
  }

  // ファイルがアップロードされているか確認
  if (!req.files || Object.keys(req.files).length === 0) {
    res.status(400).send('ファイルがアップロードされていません。');
    return;
  }

  // アップロードされたファイルがPowerPointファイルか確認
  let file = req.files.file;
  const extension = path.extname(file.name).toLowerCase();
  if (extension !== '.pptx' && extension !== '.ppsx') {
    res.status(400).send('.pptx または .ppsx 形式のファイルのみに対応しています。ファイル形式を確認してください。');
    return;
  }

  // 変換の場合は、.ppsxファイルであることを確認
  if (isConvertToPPTX && extension !== '.ppsx') {
    res.status(400).send('PPSX→PPTX変換は .ppsx 形式のファイルのみ対応しています。ファイル形式を確認してください。');
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
        .send(
          'PowerPointファイルのデータが破損しています。presentation.xml が見つかりませんでした。',
        );
      return;
    }

    // p:modifyVerifierを削除して、presentation.xmlを更新
    const xmlString = await presentationEntry.async('string');
    const newXmlString = remove_p_modifyVerifier(xmlString);
    if (newXmlString === null) {
      res
        .status(400)
        .send(
          'アップロードされたPowerPointはパスワード保護されていない可能性があります。ファイルを確認してください。',
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
        .send(
          'PowerPointファイルのデータが破損しています。[Content_Types].xml が見つかりませんでした。',
        );
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
