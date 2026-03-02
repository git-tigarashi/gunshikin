# 軍資金作成資料

AIエージェントを活用した軍資金（収益）作成のための資料・ツール集。

## 概要

このプロジェクトは、[TAISUN v2](https://github.com/taiyousan15/taisun_agent) をサブモジュールとして統合し、
収益化・資金調達のための調査・分析・コンテンツ生成を支援します。

## 構成

```
軍資金作成資料/
├── taisun_agent/   # TAISUN v2 - AI エージェント統合プラットフォーム (submodule)
├── docs/           # 調査資料・分析レポート
├── logs/           # 作業ログ
└── README.md
```

## セットアップ

```bash
# サブモジュールを含めてクローン
git clone --recurse-submodules https://github.com/git-tigarashi/gunshikin.git

# すでにクローン済みの場合
git submodule update --init --recursive
```

## taisun_agent の利用

```bash
cd taisun_agent
npm install
cp .env.example .env  # 環境変数を設定
npm start
```

詳細は [taisun_agent/README.md](./taisun_agent/README.md) を参照。

## ログ

作業ログは `logs/` ディレクトリに記録します。
