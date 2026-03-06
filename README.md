# ClassroomArchiver
Export your Google Classroom into a browsable HTML archive.

## 説明

- 所属するクラスのお知らせ・課題・資料が、HTML形式（Webサイト形式）で PC 内にダウンロードされます。
- 投稿に添付された共有ファイルもダウンロードされます。

## 利用方法

1. main.exe をダウンロードする
2. 適当な場所に main.exe を配置し、main.exe を実行する
3. ダウンロードされるデータ量が表示されるので、問題ないことを確認して、y を入力し Enter
   - y は yes, N は no を意味する
4. 気長に待つ
5. データが main.exe と同じ階層にある `classroomArchive` フォルダに保存される。

## 注意事項

### データの保存先

PC 内のデータは `classroomArchive/YYYYMMDDHHMMSS` というフォルダに保存されます。<br>
`YYYYMMDDHHMMSS` は年月日時を表しており、例えば、2026年3月6日19時34分08秒に実行した場合、`classroomArchive/20260306193408` というフォルダに保存されます。

### 保存されないファイル

- Google Forms
- 共有されたフォルダ
- 閲覧権限がないファイル
- 存在しないファイル
- 投稿に対するコメント
- 課題内の限定公開のコメント
- ドライブファイル内のコメント

上記の内容は保存されません。ご注意ください。

### ドライブにコピーされるファイル

- Google Docs ファイル
- Google Slides ファイル
- Google Sheets ファイル

上記の内容は保存されませんが、代わりに削除予定の Google アカウントの `マイドライブ/Classroom Archive/YYYYMMDDHHMMSS` 内に、ファイルのコピーが作成されます。
HTML ファイル内でのリンク先は、コピーされたファイルのリンクになっています。<br>
※ **マイドライブはアカウント停止と同時にアクセスできなくなります。** 削除予定の Google アカウントを他の個人アカウントに移行する際、Google Drive の内容をコピーすることを忘れないようにして下さい。

