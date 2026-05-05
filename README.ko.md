# OfficeWeaver

[English](README.md) | [한국어](README.ko.md) | [日本語](README.ja.md)

OfficeWeaver는 AI 친화적인 Office 문서 자동화를 만들기 위한 작은 JavaScript 런타임입니다.

LLM이 작성한 매크로 코드와 ONLYOFFICE 같은 Office 편집기 API 사이에 위치합니다. 목표는 단순합니다. 모델은 안정적인 헬퍼 API로 문서를 편집하고, 호스트 애플리케이션은 각 작업 단계의 성공/실패 결과를 구조화해서 받을 수 있게 만드는 것입니다.

현재 OfficeWeaver는 ONLYOFFICE Spreadsheet 매크로를 우선 지원합니다. 패키지에는 RAG에 바로 색인할 수 있는 버전별 `manuals/` 문서도 함께 들어 있습니다. 호스트 앱은 이 문서를 벡터 DB에 넣어두고, LLM이 코드를 작성할 때 필요한 사용법을 검색해서 프롬프트에 넣을 수 있습니다.

## 왜 필요한가

Office 편집기들은 강력한 JavaScript API를 제공합니다. 하지만 raw API를 LLM에게 그대로 맡기면 안정성이 떨어집니다.

자주 생기는 문제는 다음과 같습니다.

- 모델이 ExcelJS, Office.js, VBA, Google Apps Script, ONLYOFFICE API를 서로 섞어서 사용합니다.
- 실제로 없는 `SetFontBold`, `GetValues`, `SetFreezePanes`, `Api.CreateChart` 같은 메서드를 만들어낼 수 있습니다.
- raw JavaScript 매크로는 보통 “스크립트 하나”로 실패합니다. 호스트 앱은 실패했다는 사실은 알지만, 어느 논리 단계까지 적용됐고 어느 단계에서 실패했는지 알기 어렵습니다.
- 수식을 넣은 경우, 모델이 수식이 제대로 계산됐는지 확인하기 위해 평가된 값을 다시 보고 싶어 하는 경우가 많습니다.
- 각 애플리케이션이 헬퍼 레이어, retry 로그, 매뉴얼, 편집기 버전별 주의사항을 매번 새로 만들어야 합니다.

OfficeWeaver는 LLM에게 더 작고 추적 가능한 API를 제공해서 이 문제를 줄입니다.

## OfficeWeaver가 해결하는 것

LLM이 모든 raw 편집기 API를 직접 호출하게 하는 대신, 다음처럼 `Sheet.*` 헬퍼를 사용하게 합니다.

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "=SUM(E2:E10)");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 14 });
Sheet.freezeRows(2);
return Sheet.done("Report header updated");
```

OfficeWeaver는 이 호출들을 실제 편집기 API로 실행하고, 구조화된 결과를 반환합니다.

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

중간에 실패하면 OfficeWeaver는 즉시 중단하고 실패한 단계를 알려줍니다.

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

이 정보가 있으면 LLM은 막연히 다시 시도하지 않고, 어떤 작업이 성공했고 어디서 실패했는지 보고 다음 코드를 더 정확히 작성할 수 있습니다.

## 전체 흐름

```text
사용자 요청
  -> 호스트 앱이 LLM에게 매크로 작성을 요청
  -> LLM이 Sheet.* 코드를 작성
  -> OfficeWeaver가 헬퍼를 실행
  -> ONLYOFFICE가 문서를 변경
  -> OfficeWeaver가 구조화된 outcomes를 반환
  -> 호스트 앱이 성공을 표시하거나 실패 정보를 LLM에게 다시 전달
```

OfficeWeaver는 채팅 시스템도, 벡터 DB도, 문서 서버도 아닙니다. AI 에이전트가 Office 편집기 자동화를 더 안정적으로 수행하도록 돕는 런타임과 매뉴얼 패키지입니다.

## 설치

패키지가 npm에 배포된 뒤에는 다음처럼 설치합니다.

```bash
npm install @mythosia/officeweaver
```

배포 전 로컬 개발에서는 다음처럼 설치할 수 있습니다.

```bash
npm install ../OfficeWeaver
```

패키지 구성은 다음과 같습니다.

```text
src/officeweaver.js       편집기 플러그인 안에서 로드되는 런타임
src/officeweaver.d.ts     호스트 도구를 위한 TypeScript 선언
manuals/                  RAG 색인용 버전별 매뉴얼
README.md                 영어 가이드
README.ko.md              한국어 가이드
README.ja.md              일본어 가이드
```

## 편집기 플러그인에서 로드하기

ONLYOFFICE 플러그인에서는 LLM이 생성한 매크로를 실행하는 코드보다 먼저 OfficeWeaver를 로드합니다.

```html
<script src="plugins.js"></script>
<script src="officeweaver.js"></script>
<script src="code.js"></script>
```

호스트 앱은 보통 다음 파일을 복사합니다.

```text
node_modules/@mythosia/officeweaver/src/officeweaver.js
```

예를 들어 다음 같은 플러그인 공개 폴더로 복사할 수 있습니다.

```text
wwwroot/onlyoffice-plugin/officeweaver.js
```

OfficeWeaver는 전역 객체를 제공합니다.

```js
OfficeWeaver
```

생성된 스프레드시트 매크로는 다음 헬퍼 네임스페이스를 사용합니다.

```js
Sheet
```

## 스프레드시트 매크로 실행

호스트 앱은 LLM이 생성한 코드로 실행 가능한 명령을 만들 수 있습니다.

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

ONLYOFFICE 안에서 `command()`는 편집기의 `Api` 객체가 존재한다고 가정합니다. OfficeWeaver는 이 `Api` 객체를 내부적으로 사용합니다.

## LLM이 사용해야 하는 코드

일반적인 스프레드시트 작업에서는 LLM에게 먼저 `Sheet.*` 헬퍼를 쓰도록 지시하세요.

### 값과 수식

```js
Sheet.setValue("A1", "Sales Report");
Sheet.setFormula("F1", "=SUM(E2:E10)");
return Sheet.done("Updated title and total formula");
```

### 서식

```js
Sheet.setFill("A1:E1", "#1F4E79");
Sheet.setFont("A1:E1", { bold: true, color: "#FFFFFF", size: 12, name: "Arial" });
Sheet.setAlignment("A1:E20", { horizontal: "center", vertical: "center", wrap: true });
Sheet.setNumberFormat("E2:E20", "#,##0");
return Sheet.done("Applied table formatting");
```

### 행, 열, 레이아웃

```js
Sheet.insertRows(1, 2);
Sheet.setRowHeight(1, 30);
Sheet.setColumnWidth("A", 16);
Sheet.deleteColumns("F", 1);
return Sheet.done("Adjusted report layout");
```

### 틀 고정

```js
Sheet.freezeRows(2);
return Sheet.done("Frozen title and header rows");
```

### raw API 탈출구

모든 편집기 API가 래핑되어 있는 것은 아닙니다. 드문 작업에는 `Sheet.raw`를 사용하세요. 이렇게 하면 raw API를 쓰더라도 OfficeWeaver가 해당 작업을 outcome에 기록할 수 있습니다.

```js
Sheet.raw("customChart", { range: "A1:E10" }, function (Api, Sheet) {
  const ws = Sheet.sheet();
  return ws.AddChart("'Sheet1'!$A$1:$E$10", false, "bar", 2, 12 * 360000, 7 * 360000, 6, 0, 1, 0);
});

return Sheet.done("Added chart");
```

추적되지 않는 raw 코드보다 `Sheet.raw`를 선호하세요.

## 왜 raw ONLYOFFICE API 문서만 주면 안 되나요?

API 문서는 여전히 필요합니다. OfficeWeaver는 이 문서를 `manuals/` 아래에 제공합니다.

하지만 raw 문서만으로는 실행 추적 문제가 해결되지 않습니다.

- raw 호출은 단계별 `outcomes` 배열을 만들지 않습니다.
- 실패 시 어떤 논리 작업이 실패했는지 자동으로 알 수 없습니다.
- 수식 검증은 매크로나 호스트 앱이 매번 직접 구현해야 합니다.
- 잘못된 메서드 이름을 반복해서 교정해야 합니다.

OfficeWeaver는 두 가지를 결합합니다.

- `manuals/`는 모델에게 무엇이 존재하고 어떻게 써야 하는지 알려줍니다.
- `Sheet.*` 헬퍼는 흔한 작업을 추적 가능하고 retry하기 쉽게 만듭니다.

## RAG를 위한 manuals

패키지에는 버전별 매뉴얼이 포함되어 있습니다. 이 파일들은 평범한 Markdown 문서이지만, 호스트 앱이 벡터 DB에 색인해서 RAG로 사용할 수 있도록 메타데이터와 예제가 정리되어 있습니다.

`manuals/` 아래 파일은 임베딩 결과물이 아니라 원천 문서입니다. 호스트 애플리케이션은 이 Markdown 파일을 적절한 크기로 청킹하고, 각 청크에 대한 임베딩을 생성하고, 그 임베딩을 벡터 DB에 저장한 뒤, LLM에게 OfficeWeaver 매크로 코드를 작성시키는 시점에 관련 청크를 검색해서 프롬프트에 넣으면 됩니다.

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

호스트 앱은 설치된 패키지에서 매뉴얼을 직접 읽어야 합니다. 앱 프로젝트 안에 손으로 관리하는 복사본을 따로 만들지 않는 편이 좋습니다.

추천 메타데이터 필터:

- `engine`: `onlyoffice`
- `version_family`: `9.3`
- `kind`: `spreadsheet`

새 편집기 버전을 지원할 때는 다음처럼 새 버전 패밀리 폴더를 추가합니다.

```text
manuals/onlyoffice/9.4/spreadsheet/
```

앱이 최신 문서만 사용하고 싶다면, 이전에 임베딩한 OfficeWeaver 청크를 삭제하고 최신 manifest 기준으로 다시 색인하면 됩니다.

## 버전별 어댑터 모델

Office API는 버전에 민감합니다. ONLYOFFICE 9.3에서 작동하는 헬퍼가 ONLYOFFICE 9.4나 다른 Office 제품에서는 다른 구현을 요구할 수 있습니다.

OfficeWeaver는 다음 구조를 지향합니다.

```text
src/
  officeweaver.js
  adapters/
    onlyoffice/
      9.3/
      9.4/
    hancom/
      3.2/
      3.5/
manuals/
  onlyoffice/
    9.3/
    9.4/
  hancom/
    3.5/
```

첫 버전은 의도적으로 작게 시작하며, 현재 런타임은 하나의 파일 안에 구현되어 있습니다. 공개 설계는 추후 래퍼가 커질 때 버전별 어댑터로 나눌 수 있도록 열어둔 상태입니다.

이 구조를 사용하면 미래의 호스트 앱은 다음처럼 서로 다른 엔진 버전을 조합할 수 있습니다.

```text
ONLYOFFICE Spreadsheet: 9.3.1.2
Hancom HWP: 3.5
```

그리고 해당 엔진과 버전 패밀리에 맞는 매뉴얼만 검색하면 됩니다.

## 현재 스프레드시트 헬퍼

현재 제공되는 헬퍼는 다음과 같습니다.

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

모든 편집기 API를 첫날부터 전부 래핑하려고 하기보다, 실제 사용 중 반복적으로 실패하는 작업을 기준으로 헬퍼를 늘려가는 편이 좋습니다.

## 호스트 앱 통합 방법

일반적인 앱 통합은 네 단계입니다.

1. 패키지를 설치합니다.
2. `src/officeweaver.js`를 편집기 플러그인 빌드 출력으로 복사합니다.
3. `manuals/`를 청킹하고 임베딩해서 앱의 벡터 DB에 색인합니다.
4. 생성된 `Sheet.*` 코드를 OfficeWeaver로 실행하는 매크로 실행 도구를 등록합니다.

예시 흐름:

```text
애플리케이션 시작
  -> node_modules/@mythosia/officeweaver/manuals/manifest.json 읽기
  -> 이미 임베딩된 버전과 manifest 버전이 다르면 이전 OfficeWeaver 청크 삭제
  -> manuals/onlyoffice/9.3/spreadsheet/*.md 청킹
  -> 각 청크 임베딩 생성
  -> 벡터 DB에 저장

채팅 요청
  -> 사용자 요청과 관련된 RAG 청크 검색
  -> 검색된 매뉴얼 청크를 프롬프트에 포함
  -> LLM에게 Sheet.* 매크로 코드 작성 요청
  -> OfficeWeaver로 실행
  -> result.ok가 false이면 result.outcomes를 다시 LLM에게 보내 retry
```

## 개발

필요하면 의존성을 설치합니다.

```bash
npm install
```

테스트 실행:

```bash
npm test
```

문법 검사:

```bash
npm run check
```

npm 패키지에 어떤 파일이 들어가는지 미리 확인:

```bash
npm pack --dry-run
```

준비가 끝난 뒤 배포:

```bash
npm publish --access public
```

## 현재 한계

- 현재 런타임은 ONLYOFFICE Spreadsheet API 9.3/9.4 패밀리를 대상으로 합니다.
- 문서, 프레젠테이션, HWP 헬퍼는 아직 구현되어 있지 않습니다.
- 고급 차트 같은 복잡한 기능은 아직 `Sheet.raw`와 RAG 매뉴얼 참조가 필요할 수 있습니다.
- OfficeWeaver는 호스트 앱의 draft, 저장, 권한, undo 정책을 대체하지 않습니다.

## 이름 규칙

프로젝트 이름:

```js
OfficeWeaver
```

생성된 스프레드시트 매크로는 다음을 사용합니다.

```js
Sheet.*
```

향후 문서 타입은 다음 이름을 사용할 계획입니다.

```js
Doc.*
Slide.*
Hwp.*
```
