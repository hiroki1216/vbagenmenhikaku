Attribute VB_Name = "グローバル変数"

'給与所得に関する変数
Public annualSumS As Long '年間給与収入の格納用'
Public annualIncomeS As Long '年間給与所得格納用
Public monthlyIncomeSInput As Range '月平均給与所得出力用

'公的年金所得算出に関する変数
Public age As Long '年齢取得用'
Public annualSumP As Long '年間公的年金収入の格納用'
Public annualIncomeP As Long '年間公的年金所得格納用
Public annualIncomePInput As Range '月平均年金所得出力用
Public annualTotalIncome As Long '公的年金等に係る雑所得以外にかかる合計所得'

'所得金額調整控除額算出に関する変数
Public addDeduction As Long '所得金額調整控除加算用'

'5号減免判定処理に関する変数
Public output5 As Range '５号減免判定結果出力先'
Public outputTerm As Range '５号減免期間出力先'

'2号減免判定処理に関する変数
Public outputDecreaseRate As Range '減少率出力用'
Public outputExemptionRate As Range '減免率出力用'

'6号減免判定処理に関する変数
Public outputExePrice As Range '減免額出力用'
Public outputExeRate As Range '減免率出力用'
