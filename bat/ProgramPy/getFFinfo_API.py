import os
import tweepy

import sys
import tqdm
import numpy as np

try:
	terminal_size = os.get_terminal_size().columns
except OSError:
	terminal_size = 80

print("Twitterのフォロー、フォロワーリスト取得ツール\n")

# アクセストークンなどの識別コードを入力します
access_token= ""
access_secret = ""
api_key = ""
api_key_secret = ""

# Tweepy で Twitter API v1.1 にアクセス
auth = tweepy.OAuthHandler(api_key, api_key_secret)
auth.set_access_token(access_token, access_secret)
api = tweepy.API(auth)

# 引数からアカウントを特定
args = sys.argv
targetUser = args[1]

print("Twitterアカウント：",targetUser,"のFF数情報")
print("https://twitter.com/"+ targetUser)
# 指定したアカウントのフォローしているアカウントIDを配列に格納
follow_id_list = []
cursor = -1 # 最初の位置は-1で指定します

# すべて読み込み終わったら cursor=0 になります
while cursor!=0:
	id_cursor = api.get_friend_ids(screen_name=targetUser,cursor=cursor)
	follow_id_list += id_cursor[0] # フォロワーのIDをfollower_id_listに追加します
	cursor = id_cursor[1][1] # [previous_cursor, next_cursor]の順に格納されています

print("フォロー数　：",len(follow_id_list))

# 指定したアカウントのフォロワーのアカウントIDを配列に格納
follower_id_list = []
cursor = -1 # 最初の位置は-1で指定します

# すべて読み込み終わったら cursor=0 になります
while cursor!=0:
	id_cursor = api.get_follower_ids(screen_name=targetUser,cursor=cursor)
	follower_id_list += id_cursor[0] # フォロワーのIDをfollower_id_listに追加します
	cursor = id_cursor[1][1] # [previous_cursor, next_cursor]の順に格納されています

print("フォロワー数：",len(follower_id_list))

#フォロー数が0の場合は計算しない
if len(follow_id_list) == 0:
	print("FF比　　　　：","フォロー数が0のため、表示しません。")
else:
	print("FF比　　　　：",round(len(follower_id_list) / len(follow_id_list),2),"\n")


# ヘッダー名を埋め込む
outPutFollowList = 'ID,名前,スクリーン名,プロフィールURL,フォロー数,フォロワー数,ツイート数,追加された公開リスト数,Twitter開始日,自己紹介,リンクURL,ロケーション,鍵垢,認証済み,認証タイプ,プロフィール画像URL\n'
outPutFollowerList = 'ID,名前,スクリーン名,プロフィールURL,フォロー数,フォロワー数,ツイート数,追加された公開リスト数,Twitter開始日,自己紹介,リンクURL,ロケーション,鍵垢,認証済み,認証タイプ,プロフィール画像URL\n'

# フォローしているIDから指定の情報を取得
print("Twitterアカウント：",targetUser,"のフォロー情報を取得します。しばらくお待ち下さい…")

for i in tqdm.tqdm(range(0, len(follow_id_list), 100)):
	np.pi*np.pi
	for user in api.lookup_users(user_id=follow_id_list[i:i+100]):
		# print(user.id,',"'+user.name+'",','"@'+user.screen_name+'"','"'+user.description+'"')

		#CSV整形に影響ありそうな変数に細工する
		user_profile = user.description.replace('"', '""')
		user_location = user.location.replace('"', '""')
		user_name = user.name.replace('"', '""')

		outPutFollowList = outPutFollowList + user.id_str + ',"' + user_name + '",' + user.screen_name + ',https://twitter.com/' + user.screen_name + ',' + str(user.friends_count) + ',' + str(user.followers_count) + ',' + str(user.statuses_count) + ',' + str(user.listed_count) + ',' + str(user.created_at) + ',"' + user_profile + '",' + str(user.url) + ',"' + user_location + '",' + str(user.protected) + ',' + str(user.verified) + ',-,' + user.profile_image_url_https + '\n'

print("取得完了。\n")

# フォロワーのIDから指定の情報を取得
print("Twitterアカウント：",targetUser,"のフォロワー情報を取得します。しばらくお待ち下さい…")

for i in tqdm.tqdm(range(0, len(follower_id_list), 100)):
	np.pi*np.pi
	for user in api.lookup_users(user_id=follower_id_list[i:i+100]):
		# print(user.id,',"'+user.name+'",','"@'+user.screen_name+'"','"'+user.description+'"')

		#CSV整形に影響ありそうな変数に細工する
		user_profile = user.description.replace('"', '""')
		user_location = user.location.replace('"', '""')
		user_name = user.name.replace('"', '""')

		outPutFollowerList = outPutFollowerList + user.id_str + ',"' + user_name + '",' + user.screen_name + ',https://twitter.com/' + user.screen_name + ',' + str(user.friends_count) + ',' + str(user.followers_count) + ',' + str(user.statuses_count) + ',' + str(user.listed_count) + ',' + str(user.created_at) + ',"' + user_profile + '",' + str(user.url) + ',"' + user_location + '",' + str(user.protected) + ',' + str(user.verified) + ',-,' + user.profile_image_url_https + '\n'

print("取得完了。\n")

#フォロー、フォロワーリストファイルをUTF-8として、出力
#なお、このファイルが作れなかったら、書き込みの権限エラーか取得中でTooManyRequestエラーで止まってここまで処理が走ってない
followList = open("..\\input\\" + targetUser + '-follow.csv', 'w', encoding='utf-8')
followerList = open("..\\input\\" + targetUser + '-follower.csv', 'w', encoding='utf-8')

followList.write(outPutFollowList)
followerList.write(outPutFollowerList)

followList.close()
followerList.close()

print("FF情報をcsvとして、保存しました。\n処理は完了しました。コンソール画面を閉じて構いません。\n")