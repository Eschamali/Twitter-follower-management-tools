# Twitter-follower-management-tools
![Screenshot](docs/screenshot1.png)
Excel(VBA) と Pythonのtweepyを利用して、Twitterのフォロー、フォロワー状況を管理するツールです。<br>
※画像はプライバシー保護のため、一部青塗りしてます。

ひとまず、以下の機能を提供します。

![Screenshot](docs/screenshot2.png)
- 差分比較して、以下の状況を表示
  - フォローした
  - フォローを解除した
  - フォローされた
  - フォローを解除された

![Screenshot](docs/screenshot3.png)
- テーブルのフィルター機能で、以下の状況を表示
  - 相互
  - 片思い
  - 片思われ 

ひとまずは、上記の機能を実装しました。

# 開発経緯
ひすったーやえごったー、大手のSocialDogまでもTwitterAPIの高額な使用料の影響により<br>
フォロー、フォロワー管理のサービスの終了や機能の削減が実施されてしまいました。<br>
まだ、一部のマイナーなチェッカーサイトもありますが、いつか終わってしまうのではないかと不安になります。<br>
そこで、以下の使用条件が満たす方になりますが、簡単なフォロー、フォロワー管理をExcelで実現してみました。<br>
機能は乏しいかもしれませんが、Excelの知識があれば、もっと色んな情報が見れたりするかもしれません。<br>
ローカル管理なので、サービス都合によるデータ削除というのもありません。すべて自己管理になります。

# 使用条件
- Twitter開発者アカウントを登録している
- Twitter API v1.1が廃止される日が過ぎてもまだ生きている方<br>
- Microsoft 365 サブスクリプションに加入(Excel)
- Windows を使用

# 使用手順
## 1.ツールを使う前の準備
1. [Pythonをインストール](https://www.python.org/downloads/)
1. [tweepyをインストール](https://docs.tweepy.org/en/stable/install.html)<br>
~~~Python
pip install tweepy
~~~
## 2.Keyの編集
![Screenshot](docs/screenshot4.png)
https://developer.twitter.com/en/portal/dashboard  
bat/ProgramPy/getFFinfo_API.py　にあるファイルを開き、上記のサイトから、識別コードを入力して保存してください。

## 3.同梱のマクロ付きExcelのファイル名を編集
ファイル名をTwitterのスクリーン名<span style="color: red; ">(@から始まる文字列)</span>に変更して開いてください。<br>
これがFF情報を取得する対象アカウントになります。<br>
不本意な変更を防ぐためこのようにしました。

## 4.FF情報を取得する
![Screenshot](docs/screenshot5.png)<br>
赤丸部分を押下して、出てくる表示メッセージに問題ないか確認して、「はい」を押下するとデータ取得が始まります。

![Screenshot](docs/screenshot6.png)<br><br>

必要なソフトがインストール、設定がされていれば以下のように取得が開始されます。
![Screenshot](docs/screenshot7.png)<br>

あとは、メッセージウィンドウに従えば、データ取得とFF状況の変化を記録できると思います。

## License
[MIT License](License.txt)