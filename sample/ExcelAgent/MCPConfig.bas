' MCPブリッジとGemini APIのための設定と型定義
Option Explicit

' 設定値
Public Const MCP_BRIDGE_URL As String = "http://localhost:3001"
Public Const GEMINI_API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent"
Public Const GEMINI_API_KEY As String = "YOUR_API_KEY_HERE" ' ここにGemini APIキーを設定


' 列挙型: 会話の役割
Public Enum ConversationRole
    USER_ROLE = 1
    ASSISTANT_ROLE = 2
    TOOL_ROLE = 3
    THOUGHT_ROLE = 4 ' 思考プロセス用に追加
End Enum


' 会話履歴アイテム
Public Type ConversationItem
    role As ConversationRole
    content As String
End Type


' ツール定義構造体
Public Type ToolParameter
    name As String
    description As String
    required As Boolean
    paramType As String ' string, number, boolean など
End Type

Public Type Tool
    name As String
    description As String
    parameters() As ToolParameter
End Type


