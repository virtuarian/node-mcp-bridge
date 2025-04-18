# Node MCP Bridge - 日本語README

## 概要

Node MCP Bridgeは、Model Context Protocol（MCP）サーバーとクライアント間の調整を行うミドルウェアです。複数のMCPサーバーを一元管理し、クライアントからの要求を適切なサーバーに転送します。
主な機能には以下が含まれます：

- 複数のMCPサーバー（Playwright、Puppeteer、FileSystemなど）の管理
- RESTful APIインターフェース
- セッション管理とツール承認フロー
- Webベースの管理インターフェース
- サーバーごとのセッションタイムアウト設定（デフォルト180分、無制限も可）

## インストール

### 前提条件

- Node.js 18以上

### インストール手順

```bash
# リポジトリのクローン
git clone https://github.com/virtuarian/node-mcp-bridge.git
cd node-mcp-bridge

# 依存関係のインストール
npm install

# TypeScriptのビルド
npm run build
```

## クイックスタート

### 起動手順

```bash
# 開発モード（ソースコードの変更を監視）
npm run dev

# 本番モード
npm start
```

デフォルトでは、サーバーは`http://localhost:3001`で起動します。ポートは`.env`ファイルで変更できます：

```
PORT=8080
```


## 他のアプリからの呼び出し

Node MCP Bridgeは、RESTful APIを提供しています。他のアプリケーションからは以下のように呼び出せます：

```javascript
// セッションId(任意のID)
const sessionId = 'test-sessionid';

// ツールの呼び出し
const toolResponse = await fetch(`http://localhost:3001/tools/call/${sessionId}`, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    serverName: 'playwright',
    toolName: 'browser_navigate',
    arguments: { url: 'https://example.com' }
  })
});

const result = await toolResponse.json();
```

## API仕様

### ツール関連

| エンドポイント | メソッド | 説明 |
|--------------|--------|------|
| `/tools` | GET | 利用可能なすべてのツール一覧を取得します |
| `/tools/call` | POST | セッションなしでツールを呼び出します（自動承認ツールのみ） |
| `/tools/call/:sessionId` | POST | 特定のセッションでツールを呼び出します |
| `/tools/call/:sessionId/approve` | POST | 特定のツールの使用をセッションに承認しツールを呼び出します |
| `/tools/function-calling` | GET | LLM向けにfunction calling形式でツール一覧を取得します（クエリパラメータ: provider[必須], serverName[任意]） |
| `/tools/function-call` | POST | OpenAI形式のfunction callを実行します（リクエストボディ: name[必須], arguments[任意]） |
| `/tools/gemini-function-call` | POST | Gemini形式のfunction callを実行します（リクエストボディ: name[必須], arguments[任意]） |

### 管理API

| エンドポイント | メソッド | 説明 |
|--------------|--------|------|
| `/admin/servers` | GET | 登録されているすべてのサーバー一覧を取得します |
| `/admin/servers/:serverName` | PUT | サーバー設定を追加/更新します |
| `/admin/servers/:serverName` | DELETE | サーバーを削除します |
| `/admin/servers/:serverName/restart` | POST | サーバーを再起動します |
| `/admin/servers/:serverName/toggleDisabled` | PUT | サーバーの有効/無効状態を切り替えます |
| `/admin/servers/:serverName/tools` | GET | サーバーのツール一覧を取得します |
| `/admin/servers/:serverName/tools/:toolName/toggleAutoApprove` | PUT | ツールの自動承認設定を切り替えます |

## サーバー管理

サーバ管理画面は`http://localhost:3001/admin`でアクセスできます。

### 管理画面の使い方

管理画面では以下の操作が可能です：

1. **サーバーの追加**: 「サーバーを追加」ボタンをクリックし、必要な情報を入力します。
   - サーバー名: 一意の識別子
   - コマンド: 実行するコマンド（例: `npx`）
   - 引数: 改行区切りで指定（例: `@playwright/mcp@latest`）
   - 環境変数: 必要に応じて追加
   - タイムアウト: 操作のタイムアウト時間（秒）
   - セッションタイムアウト: セッションの有効期間（分）、0は無制限

2. **サーバーの管理**:
   - 詳細: サーバーの詳細情報を表示
   - 再起動: サーバーを再起動
   - 編集: サーバー設定を編集
   - 削除: サーバーを削除
   - 有効/無効: トグルスイッチでサーバーの状態を変更

3. **ツールの管理**:
   - サーバー詳細ページで各ツールの自動承認設定を切り替え可能

### 言語設定

右上のドロップダウンメニューから管理画面の言語を選択できます：
- English（英語）
- 日本語
- 中文（中国語）

## サンプル

### Excelからの呼び出し

ExcelマクロからNode MCP Bridgeを呼び出す例
- サンプルExcelファイルは`sample/call_excelmacro.xlsm`で確認できます。

```vb
Sub CallMcpBridge()
    Dim sessionId As String
    Dim response As String
    Dim payload As String
    Dim serverName As String
    Dim toolName As String
    Dim endPoint As String
    
    ' セッション作成
    ' 任意のセッションIDを設定します
    sessionId = "abcd"
    
    ' ツール呼び出し
    ' EndPoint
    endPoint =  "http://localhost:3001/tools/call/" & sessionId & "/approve"

    ' playwrightでhttps://example.comのページを開く
    serverName = "playwright"
    toolName = "browser_navigate"

    payload = "{""serverName"":""" & serverName & """,""toolName"":""" & toolName & """,""arguments"":{""url"":""https://example.com""}}"

    ' リクエスト送信    
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "POST", endPoint, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.send payload

    ' レスポンス取得
    response = httpRequest.responseText
    
    MsgBox "Response: " & response
End Sub
```

### Next.jsからの呼び出し

Next.jsアプリケーションからNode MCP Bridgeを呼び出す例：

```javascript
// pages/api/mcp-bridge.js
export default async function handler(req, res) {
  try {
    // 任意のセッションIDを設定（実際のアプリでは適切な識別子を使用）
    const sessionId = 'user-session-' + Math.random().toString(36).substring(2, 10);
    
    // ツールの呼び出し
    // 前の処理でbrowser_navigateしたあとにPDF保存する
    const toolRes = await fetch(`http://localhost:3001/tools/call/${sessionId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        serverName: 'playwright',
        toolName: 'browser_save_as_pdf',
        arguments: {}
      })
    });
    
    const response = await toolRes.json();
    
    // 承認が必要な場合の処理
    if (response.approvalRequired) {
      console.log('Tool requires approval, sending approval request');
      
      // 承認エンドポイントを呼び出す
      const approvalRes = await fetch(`http://localhost:3001/tools/call/${sessionId}/approve`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          serverName: response.serverName,
          toolName: response.toolName,
          arguments: {}
        })
      });
      
      const result = await approvalRes.json();
      res.status(200).json(result);
    } 
    else {
      // 自動承認またはすでに承認済みの場合
      res.status(200).json(response);
    }
  } 
  catch (error) {
    res.status(500).json({ error: error.message });
  }
}
```

フロントエンドでの使用例：

```javascript
// pages/index.js
import { useState } from 'react';

export default function Home() {
  const [pdfInfo, setPdfInfo] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  
  const savePdf = async () => {
    setLoading(true);
    setError(null);
    
    try {
      const res = await fetch('/api/mcp-bridge');
      const data = await res.json();
      
      if (data.error) {
        setError(data.error);
      } 
      else if (data.result && data.result.content) {
        // PDFの保存結果を表示（通常はテキストメッセージが返る）
        // 例: "Saved as C:\\Users\\user\\AppData\\Local\\Temp\\page-2025-04-04T07-43-22-385Z.pdf"
        const textResult = data.result.content.find(item => item.type === 'text')?.text || '';
        setPdfInfo(textResult);
      }
    } 
    catch (error) {
      console.error('Error:', error);
      setError(error.message);
    } 
    finally {
      setLoading(false);
    }
  };
  
  return (
    <div className="container mx-auto p-4">
      <h1 className="text-2xl font-bold mb-4">PDF保存</h1>
      
      <button 
        className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded"
        onClick={savePdf} 
        disabled={loading}
      >
        {loading ? '処理中...' : 'PDF保存'}
      </button>
      
      {error && (
        <div className="mt-4 p-3 bg-red-100 text-red-700 rounded">
          エラー: {error}
        </div>
      )}
      
      {pdfInfo && (
        <div className="mt-4">
          <h2 className="text-xl font-semibold mb-2">結果:</h2>
          <div className="p-4 bg-gray-100 rounded">
            <p>{pdfInfo}</p>
          </div>
        </div>
      )}
    </div>
  );
}
```


```javascript
// OpenAI API を使用したツール呼び出しの例
const response = await fetch('http://localhost:3001/tools/function-calling');
const { tools } = await response.json();

// OpenAI API に tools を渡す
const openaiResponse = await openai.chat.completions.create({
  model: "gpt-4",
  messages: [
    { role: "user", content: "ブラウザでGoogleのホームページを開いてください" }
  ],
  tools: tools
});

// OpenAI からの応答に function_call が含まれていれば実行
const message = openaiResponse.choices[0].message;
if (message.tool_calls) {
  const toolCall = message.tool_calls[0];
  
  // MCPブリッジを通じて実際のツールを実行
  const result = await fetch('http://localhost:3001/tools/function-call', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name: toolCall.function.name,
      arguments: toolCall.function.arguments
    })
  });
  
  const mcpResult = await result.json();
  console.log("実行結果:", mcpResult);
}
```


## バージョン履歴

### ver1.0.5
- Geminiのfunction callingに対応
- SampleにExcel版のAIエージェントを追加
