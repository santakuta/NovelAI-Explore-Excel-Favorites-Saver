# NovelAI Explore → Excel お気に入り保存

NovelAI の Explore（画像ギャラリー）で気に入った投稿を保存していき、1つの Excel ファイル（.xlsx）にまとめて出力する Tampermonkey 用ユーザースクリプトです。

> ⚠️ 非公式ツールです。NovelAI（Anlatan）とは関係ありません。<br>
> ⚠️ コードは全て生成AI（Claude）によって生成しています。

## 機能

- 画像のポップアップを開いて「★ Excel用に保存」ボタン（または `Alt+S`）で保存
- 保存したものを「⬇ Excel出力」でまとめて1つの .xlsx に出力
- 出力される項目：サムネイル画像、タイトル、投稿者、投稿日、プロンプト（ベース／キャラ別）、ネガティブ、ステップ数、ガイダンス、シード、サンプラー、サイズ、モデル、いいね数、URL、保存日時、メモ
- 「一覧」から保存済みの確認・削除・メモの追加
- JSON でバックアップ／復元

## インストール

1. ブラウザに [Tampermonkey](https://www.tampermonkey.net/) を入れる
2. [NovelAI-Explore-Excel-Favorites-Saver.js](./NovelAI-Explore-Excel-Favorites-Saver.js) を開き、「Raw」ボタンを押す
3. Tampermonkey のインストール画面が出るので「インストール」

## 注意事項

- 保存データはブラウザ内（Tampermonkey のストレージ）に保存されます。スクリプトを削除するとデータも消えるため、定期的に JSON バックアップをおすすめします。
- NovelAI の非公式 API を使っているため、サイトの仕様変更で動かなくなることがあります。
- 投稿作品の画像・プロンプトの権利は各投稿者にあります。出力した Excel は個人的な振り返り用とし、再配布しないでください。

## 使用ライブラリ

- [ExcelJS](https://github.com/exceljs/exceljs)（MIT License）

## ライセンス

[MIT License](./LICENSE)
