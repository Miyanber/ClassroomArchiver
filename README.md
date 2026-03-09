# ClassroomArchiver
Export your Google Classroom into a browsable HTML archive.

## 説明

- Python スクリプトです。`python main.py` で実行できます。
- Google Classroom™ で所属するクラスのお知らせ・課題・資料を HTML 形式でアーカイブすることができます。
- ここでは、「ドライブへのコピー」と「ローカルへのダウンロード」をまとめて「アーカイブ」と呼びます。

## 利用方法 (開発者向け)

1. リポジトリをクローンする
2. Google Cloud でプロジェクトを作成する
3. Google Classroom API と Google Drive API を有効にする
4. OAuth 認証情報を作成する
5. credentials.json をリポジトリ直下に作成し配置する
6. main.py を実行する
7. コンソールの内容に従って操作する
8. アーカイブが開始される
9. アーカイブが完了する

## 注意事項

### データの保存先

PC 内のデータは、実行元と同じディレクトリからの相対パスで `classroomArchive/YYYYMMDDHHMMSS` に保存されます。<br>
`YYYYMMDDHHMMSS` は年月日時を表しており、例えば、2026年3月6日19時34分08秒に実行した場合、`classroomArchive/20260306193408` というフォルダに保存されます。　　
フォルダ内にはクラスごとに HTML ファイルが生成されています。ファイル内のリンクは、ローカルにダウンロードされたものに置き換わっています。ダウンロードされていないものに関しては引き続きドライブへのリンクとなっています。

### 保存されないファイル

- Google Forms
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
- ドライブフォルダ内のファイル

上記の内容はローカルには保存されませんが、OAuth 認証を行った Google アカウントの `マイドライブ/Classroom Archive/YYYYMMDDHHMMSS` 内に、ファイルのコピーが作成されます。