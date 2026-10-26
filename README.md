# ClassroomArchiver
Export your Google Classroom into a browsable HTML archive.

## 説明

- Python スクリプトです。`python main.py` で実行できます。
- Google Classroom™ で所属するクラスのお知らせ・課題・資料を HTML 形式でアーカイブすることができます。
- 投稿の添付ファイルもアーカイブされ、ドライブにコピー或いはローカルにダウンロードされます。

## 利用方法 (開発者向け)

### 動作環境

- Python 3.10 以上
- Google アカウント（Google Classroom を利用しているもの）

### セットアップ

1. リポジトリをクローンする
2. 必要なライブラリをインストールする

```bash
pip install -r requirements.txt
```

3. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成し、Google Classroom API と Google Drive API を有効にする
4. 「OAuth 同意画面」を設定する（公開ステータスが「テスト」の場合は、利用する Google アカウントを「テストユーザー」に追加する）
5. 「認証情報」から **デスクトップアプリ** として OAuth クライアント ID を発行し、ダウンロードした JSON を `credentials.json` という名前でリポジトリ直下に配置する（形式は [credentials.json.example](./credentials.json.example) を参照）

> [!WARNING]
> `credentials.json` は `.gitignore` 済みです。絶対にコミットしないでください。

6. main.py を実行する

```bash
python ./main.py
```

7. コンソールの内容に従いブラウザで認証を完了させると、アーカイブが開始されます

### exe のビルド (Windows)

```bash
pip install -r requirements-dev.txt
python build_release.py
```

`ClassroomArchiver.zip` が生成されます。zip には `credentials.json` が同梱されるため、配布範囲に注意してください。

## 注意事項

### データの保存先

#### ローカル

ローカルの `classroomArchive` フォルダの階層は以下のようになっています。
HTMLファイル等が保存されるのはこちらです。

```
classroomArchive/
└── 20260319160909/ ←アーカイブ生成日時を表す
    ├── クラス_クラス名.html ←これがアーカイブ本体のHTMLファイル(クラスごとに生成)
    ├── driveFiles/
    │   └── id_ファイルID_name_ファイル名.拡張子 ←ダウンロードされた添付ファイル
    ├── img/
    │   ...
    ...
```

PC 内のデータは、実行元と同じディレクトリからの相対パスで `classroomArchive/YYYYMMDDHHMMSS` に保存されます。
`YYYYMMDDHHMMSS` は年月日時を表しており、例えば、2026年3月6日19時34分08秒に実行した場合、`classroomArchive/20260306193408` というフォルダに保存されます。
フォルダ内にはクラスごとに HTML ファイルが生成されています。
<br>
HTMLファイル内の添付ファイルへのリンクは、ダウンロードされたものはローカルへのリンク、ドライブにコピーされたものは、コピー元のドライブファイルへのリンクになっています。

#### ドライブ

ドライブへコピーされた添付ファイル（共有ファイルやドライブフォルダ）は以下のような階層で保存されます。

```
マイドライブ
└── Classroom Archive
    └── 20260319160909/ ←アーカイブ生成日時を表す
        └── クラス名/
            ├── 共有されたドライブフォルダ/
            │    └── フォルダ内の共有ファイル
            └── 共有ファイル
                ...

```

一部このようにドライブにコピーされるファイルも存在します。
<br>
これに関しては、Google Workspace 系のファイルはそのままの形式でダウンロードができないためです。形式変換等によるレイアウト崩れを防ぐため、ドライブへのコピーという方法を選んでいます。
<br>
また、共有フォルダに関しても同様、Google Workspace 系のファイルが含まれる可能性があるため、フォルダごとドライブにコピーします。

### 保存されない内容

- 編集権限がない Google Forms
- Google Forms の回答
- 閲覧権限がないファイル
- 存在しないファイル
- 投稿に対するコメント
- 課題内の限定公開のコメント
- ドライブファイル内のコメント

上記の内容は保存されません。ご注意ください。

### ドライブにコピーされる添付ファイル

- Google Docs ファイル
- Google Slides ファイル
- Google Sheets ファイル
- ドライブフォルダ内のファイル

上記の内容はローカルには保存されませんが、OAuth 認証を行った Google アカウントの `マイドライブ/Classroom Archive/YYYYMMDDHHMMSS` 内に、ファイルのコピーが作成されます。