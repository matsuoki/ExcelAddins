Attribute VB_Name = "mod_WorksneetFiunvtions"
Option Explicit

'Excelワークシート関数：アルファベット読みにくいやつをカナにするやつ
Function kana(r As Range) As String

    Const deli As String = "・"

    Dim i As Integer
    Dim p As String
    Dim p_all As String
    
    Dim s As String
    
    For i = 0 To Len(r.Text) - 1
        s = Mid(r.Text, i + 1, 1)
        
        If s = "!" Then p = "エクスクラメーションマーク"
        If s = """" Then p = "ダブルクオート"
        If s = "#" Then p = "シャープ"
        If s = "$" Then p = "ドル"
        If s = "%" Then p = "パーセント"
        If s = "&" Then p = "アンパサンド"
        If s = "'" Then p = "シングルクォート"
        If s = "(" Then p = "かっこ"
        If s = ")" Then p = "かっことじ"
        If s = "*" Then p = "アスタリスク"
        If s = "+" Then p = "プラス"
        If s = "," Then p = "カンマ"
        If s = "-" Then p = "ハイフン"
        If s = "." Then p = "ピリオド"
        If s = "/" Then p = "スラッシュ"
        If s = "0" Then p = "ゼロ"
        If s = "1" Then p = "イチ"
        If s = "2" Then p = "ニ"
        If s = "3" Then p = "サン"
        If s = "4" Then p = "シ"
        If s = "5" Then p = "ゴ"
        If s = "6" Then p = "ロク"
        If s = "7" Then p = "ナナ"
        If s = "8" Then p = "ハチ"
        If s = "9" Then p = "キュウ"
        If s = ":" Then p = "コロン"
        If s = ";" Then p = "セミコロン"
        If s = "<" Then p = "小なり"
        If s = "=" Then p = "イコール"
        If s = ">" Then p = "ダイナリ"
        If s = "?" Then p = "ハテナ"
        If s = "@" Then p = "アットマーク"
        If s = "A" Then p = "エー"
        If s = "B" Then p = "ビー"
        If s = "C" Then p = "シー"
        If s = "D" Then p = "デー"
        If s = "E" Then p = "イー"
        If s = "F" Then p = "エフ"
        If s = "G" Then p = "ジー"
        If s = "H" Then p = "エイチ"
        If s = "I" Then p = "アイ"
        If s = "J" Then p = "ジェイ"
        If s = "K" Then p = "ケイ"
        If s = "L" Then p = "エル"
        If s = "M" Then p = "エム"
        If s = "N" Then p = "エヌ"
        If s = "O" Then p = "オー"
        If s = "P" Then p = "ピー"
        If s = "Q" Then p = "キュー"
        If s = "R" Then p = "アール"
        If s = "S" Then p = "エス"
        If s = "T" Then p = "ティー"
        If s = "U" Then p = "ユー"
        If s = "V" Then p = "ブイ"
        If s = "W" Then p = "ダブリュー"
        If s = "X" Then p = "エックス"
        If s = "Y" Then p = "ワイ"
        If s = "Z" Then p = "ゼット"
        If s = "[" Then p = "角括弧はじめ"
        If s = "\" Then p = "円マーク"
        If s = "]" Then p = "角括弧閉じ"
        If s = "^" Then p = "キャレット"
        If s = "_" Then p = "アンダーバー"
        If s = "`" Then p = "バッククォート"
        If s = "a" Then p = "エー"
        If s = "b" Then p = "ビー"
        If s = "c" Then p = "シー"
        If s = "d" Then p = "デー"
        If s = "e" Then p = "イー"
        If s = "f" Then p = "エフ"
        If s = "g" Then p = "ジー"
        If s = "h" Then p = "エイチ"
        If s = "i" Then p = "アイ"
        If s = "j" Then p = "ジェイ"
        If s = "k" Then p = "ケイ"
        If s = "l" Then p = "エル"
        If s = "m" Then p = "エム"
        If s = "n" Then p = "エヌ"
        If s = "o" Then p = "オー"
        If s = "p" Then p = "ピー"
        If s = "q" Then p = "キュー"
        If s = "r" Then p = "アール"
        If s = "s" Then p = "エス"
        If s = "t" Then p = "ティー"
        If s = "u" Then p = "ユー"
        If s = "v" Then p = "ブイ"
        If s = "w" Then p = "ダブリュー"
        If s = "x" Then p = "エックス"
        If s = "y" Then p = "ワイ"
        If s = "z" Then p = "ゼット"
        If s = "{" Then p = "ブレースはじめ"
        If s = "|" Then p = "バー"
        If s = "}" Then p = "ブレース閉じ"
        If s = "~" Then p = "チルダ"
        
        If p = "" Then p = "（不明）"
        
        p_all = p_all + p
        
        If Not i = Len(r.Text) - 1 Then p_all = p_all + deli
    
    
    Next i
    
    kana = p_all

End Function
