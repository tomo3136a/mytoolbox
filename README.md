# mytoolbox

mytoolbox is my toolbox.
toolbox include excel addins.

- 環境管理
  - [setr](#setr) set バッチファイル作成ツール
  - [files](#files) ファイル一覧取得ツール
  - [indexed](#indexed) ファイル名インデックス設定ツール
  - [mkfolder](#mkfolder) フォルダ作成ツール
- Excel アドイン
  - [addins/myworks](#myworks) Excel 作業補助ツール
  - [addins/mydesigner](#mydesigner) Excel 図形描画ツール
  - [addins/AddinDev](#addindev) Excel マクロアドイン操作ツール

### build

事前設定：  
Excel を起動し、`オプション`-`トラストセンター`-`トラストセンターの設定`-`マクロの設定`の
「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を一時的にチェックします。ビルド終了後はチェックを外します。

```shellsession
$ git clone https://github.com/tomo3136a/mytoolbox.git mtb
$ cd mtb
$ build.cmd
```

### install

```shellsession
$ cd package
$ mtb_*.exe
$ cd mtb
$ install.cmd
$ install-addins.cmd
```

Excel アドイン(アドイン開発用)

```shellsession
$ install-addins.cmd dev
```

## setr

set バッチファイル作成ツール

## files

ファイル一覧取得ツール

## indexed

ファイル名インデックス設定ツール

## mkfolder

フォルダ作成ツール

## myworks

Excel 作業補助ツール

## mydesigner

Excel 図形描画ツール

## AddinDev

Excel マクロアドイン操作ツール

##
