# Kotobuku Auto Notify DEMO20

## 概要
Google Apps Script によるステータス自動通知デモ。
スプレッドシートの状態変化をトリガーに、
工事完了通知メールを自動送信する。

## 構成
- Google Spreadsheet
- Google Apps Script
- Gmail

## 主な機能
- 送信対象行の判定
- テンプレート差し込み
- dry-run / 本番送信切り替え
- 送信履歴・結果の記録

## ファイル構成
- demo20_mailer.gs : メイン処理
- demo20_lib.gs : 共通ユーティリティ
- send_control_core.gs : 制御ロジック
- reset.gs : デモ用リセット（本番未使用）

## 注意
本リポジトリはレビュー・設計確認用。
実行環境とは分離して管理する。
