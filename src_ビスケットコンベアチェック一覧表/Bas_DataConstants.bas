Attribute VB_Name = "Bas_DataConstants"
Option Explicit

' --- ADODB.Connection / Recordset State ---

'/** @brief オブジェクトが開いている状態を示す定数 */
Public Const AD_STATE_OPEN As Long = 1

' --- ADODB.StreamTypeEnum ---

'/** @brief ADODB.Stream のテキストモードを示す定数 */
Public Const AD_TYPE_TEXT As Long = 2

' --- ADODB.CursorLocation ---

'/** @brief クライアントサイドカーソルを使用することを示す定数 */
Public Const AD_USE_CLIENT As Long = 3

' --- ADODB.CursorType ---

'/** @brief 静的カーソル (データの読み取りに最適) を示す定数 */
Public Const AD_OPEN_STATIC As Long = 3

' --- ADODB.LockType ---

'/** @brief 読み取り専用ロックを示す定数 */
Public Const AD_LOCK_READ_ONLY As Long = 1

' --- ADODB.CommandTypeEnum ---
'/** @brief テキストコマンド (SQL文など) を示す定数 */
Public Const AD_CMD_TEXT As Long = 1

' --- ADODB.ParameterDirectionEnum ---
'/** @brief 入力パラメータを示す定数 */
Public Const AD_PARAM_INPUT As Long = 1

' --- ADODB.DataTypeEnum ---
'/** @brief 整数型 (Long) を示す定数 */
Public Const AD_INTEGER As Long = 3
'/** @brief 可変長文字列型 (String) を示す定数 */
Public Const AD_VAR_CHAR As Long = 200

