# Electron を用いたデスクトップアプリ制作

## はじめに

Electron と Python (Flask) を使ったデスクトップアプリの制作についてまとめる

## 開発環境

- Windows 11
- Python 3.10.4

## 環境構築

- Node.js
- Electron
- Flask

## Electron

### package.json の作成

作業ディレクトリを作成し、そこで npm init する

- -y : 初期値で package.json を生成

```sh
mkdir sample_app
cd my_electron
npm init -y
```

- package.json

```json
{
  "name": "sample_app",
  "version": "1.0.0",
  "description": "",
  "main": "index.js",
  "scripts": {
    "test": "echo \"Error: no test specified\" && exit 1"
  },
  "keywords": [],
  "author": "",
  "license": "ISC"
}
```

## 参考資料

- [ElectronとPython（Flask）を使ったデスクトップアプリの制作からパッケージングまで](https://qiita.com/goto_y/items/6cabe72da415755b29b5)