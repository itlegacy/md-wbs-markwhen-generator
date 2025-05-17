---
title: ServiceDev 新サービス開発プロジェクト (v6)
description: 次世代向け新サービスの開発プロジェクト。
date: 2025-05-16
projectstartdate: 2025-06-01
projectoveralldeadline: 2026-04-15
view: month
---

## 1. 企画・調査フェーズ
    - deadline: 2025-08-31 # このセクションの目標完了日
    - status: inprogress
    - assignee: 企画チーム

このフェーズでは、市場の動向を調査し、新サービスの基本的な企画を固めます。
ユーザーニーズの把握と競合との差別化が重要となります。

### 1.1. 市場調査グループ
    - deadline: 2025-07-25 # このグループの目標完了日
    - status: completed
    - assignee: 調査チーム

#### 1.1.1. 市場全体のトレンド分析
    - duration: 10d
    - deadline: 2025-06-15
    - status: completed
    - progress: 100%

最新の技術トレンドと消費者行動の変化を調査します。

#### 1.1.2. ターゲット顧客セグメント定義
    - duration: 7d
    - deadline: 2025-06-30
    - depends: 1.1.1
    - status: completed
    - progress: 100%

#### 1.1.3. 主要プレイヤーと提供価値の分析
    - duration: 8d
    - deadline: 2025-07-25 # MD-WBS修正: 1.1.3のdeadlineを先行タスク1.1.2の終了(2025-07-08)より後に
    - depends: 1.1.2
    - status: inprogress
    - progress: 60%

### 1.2. 競合分析グループ
    - deadline: 2025-08-08
    - status: inprogress
    - assignee: 調査チーム

#### 1.2.1. 競合製品Aの機能・価格調査
    - duration: 5d
    - deadline: 2025-08-01 # MD-WBS修正: 先行1.1.3の終了(2025-07-25)より後
    - depends: 1.1.3
    - status: inprogress
    - progress: 30%

#### 1.2.2. 競合製品Bのユーザーレビュー分析
    - duration: 5d
    - deadline: 2025-08-08
    - depends: 1.2.1
    - status: pending

#### 1.2.3. 技術的優位性と劣位性の評価
    - duration: 7d
    - deadline: 2025-08-19
    - depends: 1.2.2
    - status: pending

### 1.3. 企画書作成グループ
    - deadline: 2025-09-12
    - status: pending
    - assignee: 企画リード

#### 1.3.1. 企画骨子作成とレビュー
    - duration: 10d
    - deadline: 2025-08-30
    - depends: 1.2.3
    - status: pending

#### 1.3.2. 企画書詳細作成
    - duration: 10d
    - deadline: 2025-09-12
    - depends: 1.3.1
    - status: pending


## 2. 設計フェーズ
    - deadline: 2025-11-30
    - depends: 1. # セクション1全体に依存

### 2.1. 要件定義グループ
    - deadline: 2025-10-27 # MD-WBS修正: 2.1.2の終了日に合わせる
    - assignee: PM, 設計チーム

#### 2.1.1. ビジネス要件定義
    - duration: 10d
    - deadline: 2025-09-29 # MD-WBS修正: 先行1.3.2の終了(2025-09-12)より後、かつサブタスクを包含できるように
    - depends: 1.3.2
    - status: pending

このタスクでは、ビジネス上のゴールと主要な機能を定義します。
ステークホルダーへのヒアリングが重要です。

##### 2.1.1.1. ユーザーインタビュー実施 (3回)
    - duration: 5d
    - deadline: 2025-09-19 # MD-WBS修正: 親2.1.1のdeadlineより前、かつ先行1.3.2の終了より後
    - depends: 1.3.2
    - status: pending

##### 2.1.1.2. 要求仕様書作成
    - duration: 5d
    - deadline: 2025-09-27 # MD-WBS修正: 親2.1.1のdeadlineより前、かつ先行2.1.1.1の終了より後
    - depends: 2.1.1.1
    - status: pending

#### 2.1.2. システム要件定義
    - duration: 10d
    - deadline: 2025-10-14 # MD-WBS修正: 先行2.1.1の終了(2025-10-10)より後
    - depends: 2.1.1
    - status: pending

## 3. 依存関係テストセクション
    - deadline: 2026-04-15

### 3.1. 正常な依存関係テストグループ
    - deadline: 2026-03-31

#### 3.1.1. 先行タスクA (正常)
    - duration: 5d
    - deadline: 2026-03-10
    - status: pending

#### 3.1.2. 後続タスクB (正常)
    - duration: 5d
    - deadline: 2026-03-20
    - depends: 3.1.1
    - status: pending

### 3.2. 意図的な矛盾テストグループ
    - deadline: 2026-04-15

#### 3.2.1. 先行タスクX (矛盾テスト用)
    - duration: 7d
    - deadline: 2026-04-05
    - status: pending

#### 3.2.2. 後続タスクY (意図的に矛盾発生)
    - description: このタスクは先行タスクXの終了日より前に開始するようにdeadlineを設定しているため、警告が出るはず。
    - duration: 3d
    - deadline: 2026-04-01
    - depends: 3.2.1
    - status: pending