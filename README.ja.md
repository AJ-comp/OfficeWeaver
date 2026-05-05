# OfficeWeaver

[English](README.md) | [한국어](README.ko.md) | [日本語](README.ja.md)

OfficeWeaver は、AI フレンドリーな Office 文書自動化を作るための小さな JavaScript ランタイムです。

LLM が生成したマクロコードと、ONLYOFFICE などの Office エディター API の間に入ります。目的はシンプルです。モデルには安定したヘルパー API で文書を編集させ、ホストアプリケーションには各ステップの成功・失敗を構造化された結果として返します。

現在の OfficeWeaver は、ONLYOFFICE Spreadsheet マクロを中心にしています。パッケージには、RAG でそのままインデックスできるバージョン別の `manuals/` も含まれています。ホストアプリはこれらの文書をベクトル DB に入れておき、LLM がコードを書くときに必要な使い方を検索してプロンプトに渡せます。

## なぜ必要なのか

Office エディターは強力な JavaScript API を提供しています。しかし raw API をそのまま LLM に使わせると、信頼性の問題が出やすくなります。

よくある問題:

- モデルが ExcelJS、Office.js、VBA、Google Apps Script、ONLYOFFICE API を混同します。
- 実際には存在しない `SetFontBold`、`GetValues`、`SetFreezePanes`、`Api.CreateChart` のようなメソッド名を作ってしまうことがあります。
- raw JavaScript マクロは、多くの場合「ひとつの大きなスクリプト」として失敗します。ホストアプリは失敗したことは分かりますが、どの論理ステップまで適用され、どこで失敗したのかを把握しにくいです。
- 数式を書き込む場合、モデルは数式が正しく計算されたか確認するために評価後の値を見たくなることがあります。
- 各アプリケーションが、ヘルパーレイヤー、retry ログ、マニュアル、エディターのバージョン別注意点を毎回作ることになります。

OfficeWeaver は、LLM により小さく、追跡可能な API を提供することでこれらの問題を減らします。

## OfficeWeaver が解決すること

LLM にすべての raw エディター API を直接呼ばせる代わりに、次のような `Sheet.*` ヘルパーを生成させます。

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "=SUM(E2:E10)");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 14 });
Sheet.freezeRows(2);
return Sheet.done("Report header updated");
```

OfficeWeaver はこれらの呼び出しを実際のエディター API で実行し、構造化された結果を返します。

```json
{
  "ok": true,
  "engine": "onlyoffice",
  "version": "9.3.1.2",
  "version_family": "9.3",
  "summary": "Report header updated",
  "applied": 4,
  "failed": 0,
  "outcomes": [
    { "step": 1, "action": "setValue", "address": "A1", "ok": true },
    { "step": 2, "action": "setFormula", "address": "F1", "ok": true, "evaluated_value": 123000 },
    { "step": 3, "action": "setFont", "address": "A1:E1", "ok": true },
    { "step": 4, "action": "freezeRows", "rows": 2, "ok": true }
  ]
}
```

途中で失敗した場合、OfficeWeaver はすぐに停止し、失敗したステップを返します。

```json
{
  "ok": false,
  "summary": "Macro failed at step 3",
  "applied": 2,
  "failed": 1,
  "stopped_at": 3,
  "outcomes": [
    { "step": 1, "action": "setValue", "address": "A1", "ok": true },
    { "step": 2, "action": "setFont", "address": "A1:E1", "ok": true },
    { "step": 3, "action": "insertRows", "at": 1, "ok": false, "error": "..." }
  ]
}
```

この情報があれば、LLM はやみくもに再試行するのではなく、何が成功してどこで失敗したのかを見て、次のコードをより正確に書けます。

## 全体像

```text
ユーザーの要求
  -> ホストアプリが LLM にマクロ作成を依頼
  -> LLM が Sheet.* コードを書く
  -> OfficeWeaver がヘルパーを実行
  -> ONLYOFFICE が文書を変更
  -> OfficeWeaver が構造化された outcomes を返す
  -> ホストアプリが成功を表示するか、失敗情報を LLM に返す
```

OfficeWeaver はチャットシステムでも、ベクトル DB でも、文書サーバーでもありません。AI エージェントが Office エディター自動化をより安定して行うためのランタイムとマニュアルパッケージです。

## インストール

npm からインストールします。

```bash
npm install @mythosia/officeweaver
```

パッケージには次のものが含まれています。

```text
src/officeweaver.js       エディタープラグイン内で読み込むランタイム
src/officeweaver.d.ts     ホストツール向け TypeScript 宣言
manuals/                  RAG インデックス用のバージョン別マニュアル
README.md                 英語ガイド
README.ko.md              韓国語ガイド
README.ja.md              日本語ガイド
```

## エディタープラグインで読み込む

ONLYOFFICE プラグインでは、LLM が生成したマクロを実行するコードより前に OfficeWeaver を読み込みます。

```html
<script src="plugins.js"></script>
<script src="officeweaver.js"></script>
<script src="code.js"></script>
```

ホストアプリは通常、次のファイルをコピーします。

```text
node_modules/@mythosia/officeweaver/src/officeweaver.js
```

たとえば次のようなプラグイン公開フォルダーへコピーします。

```text
wwwroot/onlyoffice-plugin/officeweaver.js
```

OfficeWeaver はグローバルオブジェクトを公開します。

```js
OfficeWeaver
```

生成されたスプレッドシートマクロは、次のヘルパーネームスペースを使います。

```js
Sheet
```

## スプレッドシートマクロを実行する

ホストアプリは、LLM が生成したコードから実行可能なコマンドを作れます。

```js
const command = OfficeWeaver.buildSpreadsheetMacroCommand(`
Sheet.setValue("A1", "Sales Report");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 14 });
Sheet.setFill("A1:E1", "#1F4E79");
return Sheet.done("Styled report title");
`, {
  engine: "onlyoffice",
  version: "9.3.1.2"
});

const result = command();
```

ONLYOFFICE 内では、`command()` はエディターの `Api` オブジェクトが存在することを前提にします。OfficeWeaver はその `Api` オブジェクトを内部で使います。

## LLM が使うべきコード

一般的なスプレッドシート作業では、LLM にまず `Sheet.*` ヘルパーを使わせてください。

### 値と数式

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "=SUM(E2:E10)");
return Sheet.done("Updated title and total formula");
```

### 書式

```js
Sheet.setFill("A1:E1", "#1F4E79");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 12, name: "Arial" });
Sheet.setAlignment("A1:E20", { horizontal: "center", vertical: "center", wrap: true });
Sheet.setNumberFormat("E2:E20", "#,##0");
return Sheet.done("Applied table formatting");
```

### 行、列、レイアウト

```js
Sheet.insertRows(1, 2);
Sheet.setRowHeight(1, 30);
Sheet.setColumnWidth("A", 16);
Sheet.deleteColumns("F", 1);
return Sheet.done("Adjusted report layout");
```

### ウィンドウ枠の固定

```js
Sheet.freezeRows(2);
return Sheet.done("Frozen title and header rows");
```

### raw API の逃げ道

すべてのエディター API がラップされているわけではありません。まれな操作には `Sheet.raw` を使います。これにより raw API を使っても OfficeWeaver がその操作を outcome に記録できます。

```js
Sheet.raw("customChart", { range: "A1:E10" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  return ws.AddChart("'Sheet1'!$A$1:$E$10", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
});

return Sheet.done("Added chart");
```

追跡されない raw コードよりも `Sheet.raw` を優先してください。

## raw ONLYOFFICE API ドキュメントだけではだめなのか

API ドキュメントは必要です。OfficeWeaver はそれを `manuals/` 以下に同梱しています。

ただし raw ドキュメントだけでは、実行の観測性は解決できません。

- raw 呼び出しはステップごとの `outcomes` 配列を作りません。
- 失敗時にどの論理操作が失敗したかを自動では分かりません。
- 数式検証は、マクロやホストアプリが毎回実装する必要があります。
- 間違ったメソッド名を繰り返し修正する必要があります。

OfficeWeaver は次の二つを組み合わせます。

- `manuals/` は、何が存在し、どう使うべきかをモデルに教えます。
- `Sheet.*` ヘルパーは、よく使う操作を追跡可能にし、retry しやすくします。

## RAG のための manuals

パッケージにはバージョン別のマニュアルが含まれています。これらは普通の Markdown ファイルですが、ホストアプリがベクトル DB にインデックスして RAG で使えるように、メタデータと例が整理されています。

`manuals/` 配下のファイルは埋め込み結果ではなく、元の文書です。ホストアプリケーションはこれらの Markdown ファイルを適切なサイズにチャンク化し、各チャンクの embedding を作成し、その embedding をベクトル DB に保存します。そして LLM に OfficeWeaver マクロコードを書かせるタイミングで、関連するチャンクを検索してプロンプトに含めます。

```text
manuals/
  manifest.json
  onlyoffice/
    9.3/
      spreadsheet/
        OfficeWeaver.SheetHelpers.md
        ApiWorksheet.AddChart.md
        examples.ModernTable.md
        errors.CommonWrongApis.md
```

ホストアプリは、インストール済みパッケージからマニュアルを直接読むべきです。アプリプロジェクト内に手動管理のコピーを作らない方が安全です。

推奨メタデータフィルター:

- `engine`: `onlyoffice`
- `version_family`: `9.3`
- `kind`: `spreadsheet`

OfficeWeaver が別のエディターバージョンを含む場合、そのバージョンは次のように独立したバージョンファミリーフォルダーとして提供されます。

```text
manuals/onlyoffice/9.4/spreadsheet/
```

アプリが最新の文書だけを使いたい場合は、以前に埋め込んだ OfficeWeaver チャンクを削除し、最新の manifest に基づいて再インデックスします。

## バージョンを合わせる

Office API はバージョンに敏感です。ONLYOFFICE 9.3 で動くヘルパーが、ONLYOFFICE 9.4 や別の Office 製品では別実装を必要とすることがあります。

マクロを生成または実行するときは、実際に実行しているエディターのエンジンとバージョンを渡してください。

```js
OfficeWeaver.buildSpreadsheetMacroCommand(code, {
  engine: "onlyoffice",
  version: "9.3.1.2"
});
```

manuals をインデックスし検索するときも、同じエンジンとバージョンファミリーをフィルターとして使ってください。

```text
engine: onlyoffice
version_family: 9.3
kind: spreadsheet
```

これにより、異なるエディターバージョンの API 説明が混ざることを避けられます。

## 現在のスプレッドシートヘルパー

現在提供されているヘルパーは次の通りです。

- `Sheet.setValue(address, value, sheetName?)`
- `Sheet.setFormula(address, formula, sheetName?)`
- `Sheet.setNumberFormat(address, format, sheetName?)`
- `Sheet.setFont(address, options, sheetName?)`
- `Sheet.setFill(address, color, sheetName?)`
- `Sheet.setBorder(address, options, sheetName?)`
- `Sheet.setAlignment(address, options, sheetName?)`
- `Sheet.merge(address, sheetName?, across?)`
- `Sheet.unmerge(address, sheetName?)`
- `Sheet.insertRows(at, count?, sheetName?)`
- `Sheet.deleteRows(at, count?, sheetName?)`
- `Sheet.insertColumns(at, count?, sheetName?)`
- `Sheet.deleteColumns(at, count?, sheetName?)`
- `Sheet.setColumnWidth(column, width, sheetName?)`
- `Sheet.setRowHeight(row, height, sheetName?)`
- `Sheet.addSheet(name)`
- `Sheet.renameSheet(from, to)`
- `Sheet.deleteSheet(name)`
- `Sheet.setActiveSheet(name)`
- `Sheet.freezeRows(rows, sheetName?)`
- `Sheet.freezeColumns(columns, sheetName?)`
- `Sheet.freezeAt(address, sheetName?)`
- `Sheet.unfreeze(sheetName?)`
- `Sheet.raw(action, details, fn)`
- `Sheet.outcomes()`
- `Sheet.done(summary, data?)`

ここに載っていない操作は `Sheet.raw(...)` で包んで使えます。raw API を使う場合でも、実行結果は `outcomes` に残ります。

## ホストアプリへの統合方法

一般的な統合は 4 ステップです。

1. パッケージをインストールします。
2. `src/officeweaver.js` をエディタープラグインのビルド出力へコピーします。
3. `manuals/` をチャンク化し、embedding を作成して、アプリのベクトル DB にインデックスします。
4. 生成された `Sheet.*` コードを OfficeWeaver で実行するマクロ実行ツールを登録します。

例:

```text
アプリケーション起動
  -> node_modules/@mythosia/officeweaver/manuals/manifest.json を読む
  -> 既に埋め込まれたバージョンと manifest のバージョンが違えば古い OfficeWeaver チャンクを削除
  -> manuals/onlyoffice/9.3/spreadsheet/*.md をチャンク化
  -> 各チャンクの embedding を作成
  -> ベクトル DB に保存

チャット要求
  -> ユーザー要求に関連する RAG チャンクを検索
  -> 検索されたマニュアルチャンクをプロンプトに含める
  -> LLM に Sheet.* マクロコードを書かせる
  -> OfficeWeaver で実行
  -> result.ok が false なら result.outcomes を LLM に返して retry
```

## 現在の制限

- 現在のランタイムには ONLYOFFICE Spreadsheet API 9.3/9.4 ファミリー向けのスプレッドシートヘルパーが含まれています。
- Document、Presentation、HWP のヘルパーはこのパッケージにはまだ含まれていません。
- 高度なチャートなど複雑な機能は、まだ `Sheet.raw` と RAG マニュアル参照が必要になる場合があります。
- OfficeWeaver はホストアプリの draft、保存、権限、undo ポリシーを置き換えるものではありません。

## 命名

プロジェクト名:

```js
OfficeWeaver
```

生成されたスプレッドシートマクロは次を使います。

```js
Sheet.*
```
