---
title: ServiceDev 新サービス開発プロジェクト (v6.1 - org属性追加)
description: 次世代向け新サービスの開発プロジェクト。org属性テストケースを含む。
date: 2025-05-17
projectstartdate: 2025-06-01
projectoveralldeadline: 2026-04-15
view: month
---

## 1. 企画・調査フェーズ
    - deadline: 2025-08-31
    - status: inprogress
    - org: 全社プロジェクト部門 # セクションレベルの組織
    - assignee: 企画チームリーダー

このフェーズでは、市場の動向を調査し、新サービスの基本的な企画を固めます。
ユーザーニーズの把握と競合との差別化が重要となります。

### 1.1. 市場調査グループ
    - deadline: 2025-07-25
    - status: completed
    - org: マーケティング部 # グループレベルの組織
    - assignee: 調査チーム担当A

#### 1.1.1. 市場全体のトレンド分析
    - duration: 10d
    - deadline: 2025-06-15
    - status: completed
    - progress: 100%
    - org: マーケティング部 # タスクにもorgを指定可能 (グループと同じ場合が多い)
    - assignee: 佐藤

最新の技術トレンドと消費者行動の変化を調査します。

#### 1.1.2. ターゲット顧客セグメント定義
    - duration: 7d
    - deadline: 2025-06-30
    - depends: 1.1.1
    - status: completed
    - progress: 100%
    - org: マーケティング部
    - assignee: 鈴木

#### 1.1.3. 主要プレイヤーと提供価値の分析
    - duration: 8d
    - deadline: 2025-07-25
    - depends: 1.1.2
    - status: inprogress
    - progress: 60%
    - org: マーケティング部
    - assignee: 高橋

### 1.2. 競合分析グループ
    - deadline: 2025-08-08
    - status: inprogress
    - org: 事業戦略室
    - assignee: 分析チーム担当B

#### 1.2.1. 競合製品Aの機能・価格調査
    - duration: 5d
    - deadline: 2025-08-01
    - depends: 1.1.3
    - status: inprogress
    - progress: 30%
    - assignee: 田中

#### 1.2.2. 競合製品Bのユーザーレビュー分析
    - duration: 5d
    - deadline: 2025-08-08
    - depends: 1.2.1
    - status: pending
    - assignee: 伊藤

#### 1.2.3. 技術的優位性と劣位性の評価
    - duration: 7d
    - deadline: 2025-08-19
    - depends: 1.2.2
    - status: pending
    - org: 技術調査部門 # このタスクだけ別の組織が担当する場合など
    - assignee: 渡辺

### 1.3. 企画書作成グループ
    - deadline: 2025-09-12
    - status: pending
    - org: プロジェクト推進室
    - assignee: 企画リード主任

#### 1.3.1. 企画骨子作成とレビュー
    - duration: 10d
    - deadline: 2025-08-30
    - depends: 1.2.3
    - status: pending
    - assignee: 山本

#### 1.3.2. 企画書詳細作成
    - duration: 10d
    - deadline: 2025-09-12
    - depends: 1.3.1
    - status: pending
    - assignee: 中村


## 2. 設計フェーズ
    - deadline: 2025-11-30
    - depends: 1.
    - org: 開発本部

### 2.1. 要件定義グループ
    - deadline: 2025-10-27
    - org: 開発第一部
    - assignee: PMリード、設計エキスパート

#### 2.1.1. ビジネス要件定義
    - duration: 10d
    - deadline: 2025-09-29
    - depends: 1.3.2
    - status: pending
    - assignee: プロダクトオーナー

このタスクでは、ビジネス上のゴールと主要な機能を定義します。
ステークホルダーへのヒアリングが重要です。

##### 2.1.1.1. ユーザーインタビュー実施 (3回)
    - duration: 5d
    - deadline: 2025-09-19
    - depends: 1.3.2
    - status: pending
    - org: UXチーム
    - assignee: UXリサーチャー

##### 2.1.1.2. 要求仕様書作成
    - duration: 5d
    - deadline: 2025-09-27
    - depends: 2.1.1.1
    - status: pending
    - assignee: ビジネスアナリスト

#### 2.1.2. システム要件定義
    - duration: 10d
    - deadline: 2025-10-14
    - depends: 2.1.1
    - status: pending
    - org: 開発第二部
    - assignee: システムアーキテクト

## 3. 依存関係テストセクション
    - deadline: 2026-04-15
    - org: QA部門

### 3.1. 正常な依存関係テストグループ
    - deadline: 2026-03-31

#### 3.1.1. 先行タスクA (正常)
    - duration: 5d
    - deadline: 2026-03-10
    - status: pending
    - assignee: テスターA

#### 3.1.2. 後続タスクB (正常)
    - duration: 5d
    - deadline: 2026-03-20
    - depends: 3.1.1
    - status: pending
    - assignee: テスターB

### 3.2. 意図的な矛盾テストグループ
    - deadline: 2026-04-15

#### 3.2.1. 先行タスクX (矛盾テスト用)
    - duration: 7d
    - deadline: 2026-04-05
    - status: pending
    - assignee: テスターC

#### 3.2.2. 後続タスクY (意図的に矛盾発生)
    - description: このタスクは先行タスクXの終了日より前に開始するようにdeadlineを設定しているため、警告が出るはず。
    - duration: 3d
    - deadline: 2026-04-01
    - depends: 3.2.1
    - status: pending
    - assignee: テスターD