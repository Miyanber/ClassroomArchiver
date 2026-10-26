# Security Policy

## 脆弱性の報告 / Reporting a Vulnerability

セキュリティ上の問題を発見した場合は、**公開 Issue ではなく** GitHub の
[Private vulnerability reporting](https://github.com/Miyanber/ClassroomArchiver/security/advisories/new)
からご報告ください。

Please report security issues privately via GitHub's
[Private vulnerability reporting](https://github.com/Miyanber/ClassroomArchiver/security/advisories/new)
instead of opening a public issue.

## 認証情報の取り扱い

- `credentials.json`(OAuth クライアント情報)は各自の Google Cloud プロジェクトで発行し、**リポジトリにコミットしないでください**。
- 本ツールはユーザーのパスワードを扱いません。認証は Google の OAuth 2.0 フローで行われます。
