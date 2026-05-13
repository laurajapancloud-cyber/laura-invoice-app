import os
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '.env'))

from main import db_conn

raw_data = """KEN	AI1022	55	004	緒方
べにや	AK2002	55	004	緒方
アベ	ＡK2004	55	004	緒方
裕コレクションズ	AM2003	55	004	緒方
山崎洋服店	AM2004	55	004	緒方
アクティブY	AM2005	60	004	緒方
タイナカ	AM2006	55	004	緒方
オリガー	AM2007	50	004	緒方
ヴォーグサイトー	AM2010	55	004	緒方
アオキヤ	FK2001	60	004	緒方
マルモ洋服店	FK2004	50	004	緒方
田屋洋品店	GM2001	50	004	緒方
ロルム	GM2003	55	004	緒方
BEAT	GM2004	60	004	緒方
エンゼル	GM2016	60	004	緒方
多田洋品店	GM2017	55	004	緒方
Kat' club	HK2002	60	004	緒方
マルタミ	HK2003	55	004	緒方
くらいむ	HK2008	55	004	緒方
D．R．M．	HK2010	60	004	緒方
ふたば	HK2013	55	004	緒方
クラブカンパニー2nd	IW2003	55	004	緒方
トミコシ	KN2004	55	008	緒方
大　成　堂	KN2005	55	004	緒方
ミナト貿易	KN2007	55	004	緒方
Ｋプリンス	KN2009	55	004	緒方
ジェンツスポーツ	KN2012	58	004	緒方
ジャックラビッツ	KN2017	60	008	緒方
アプローズ	KN2019	58	004	緒方
ヨ　シ　ダ	MG2001	50	004	緒方
K2クラブ	MG2002	55	004	緒方
ホバラヤ	MG2006	55	004	緒方
二　葉	SI2002	55	004	緒方
シーダ　　	SI2010	50	004	緒方
ニック04	SZ2024	60	004	緒方
花　菱	SZ2003	55	004	緒方
エコー(増田）	SZ2004	55	004	緒方
タ　ム	SZ2006	55	004	緒方
ヒデ　	SZ2007	55	004	緒方
ルック青島	SZ2008	55	004	緒方
プレイスユー・オム	SZ2009	55	004	緒方
ノーマン	SZ2013	55	004	緒方
トミーゴルフ	SZ2015	55	004	緒方
セブンエビス	SZ2016	55	004	緒方
フォーシーズンギャラリー	SZ2019	55	004	緒方
エコー御殿場	SZ2020	55	004	緒方
ジラソーレ	SZ2021	50	004	緒方
ナナミ	SZ2017	55	004	緒方
㈱吉原屋洋服店	SZ2026	60	004	緒方
エコー	SZ2027	52	004	緒方
スズキ洋服店	SZ2028	55	004	緒方
エコー三島店	SZ2029	55	004	緒方
アルファ	TG2004	60	004	緒方
モトキ四丁目店	TK2046	50	004	緒方
MONACO	TK2051	60	004	緒方
田屋(事務所）	TK2057	60	004	緒方
DECENT	TK2061	55	004	緒方
スコットランド倶楽部　	YM2007	60	004	緒方
三枝	YT2001	55	004	緒方
マルハチ洋服店	YT2002	55	004	緒方
ジェルシー	YT2003	55	004	緒方
ラブ	YT2004	55	004	緒方
ミ　ラ　ノ	CB2001	55	008	髙橋
アルティジャーノ　ウォモ	CB2007	60	008	髙橋
ワカマツ	CB2012	55	008	髙橋
MORI	CB2016	55	008	髙橋
わたなべ	FK2003	55	008	髙橋
三星	GM2012	55	008	髙橋
マイノリティ	IB2001	55	008	髙橋
戸　塚	IB2004	55	008	髙橋
ナカダ	IB2005	55	008	髙橋
オーテル	IB2006	55	008	髙橋
スタッグ	IB2007	55	008	髙橋
コイケ	IB2008	55	008	髙橋
ヨシワラ	IB2009	55	008	髙橋
ノブス	ＩＢ2010	55	008	髙橋
モリアダックス	KN2006	55	008	髙橋
コメジ	NA2001	55	008	髙橋
だいこくや	NA2002	55	008	髙橋
ヤマダドレス	NA2003	60	008	髙橋
男子専科　梅村	NA2005	55	008	髙橋
マルモ	ＮＡ2007	50	008	髙橋
カナザワ	NG2001	55	008	髙橋
プラザゴルフ	NG2003	55	008	髙橋
グレンズ	NG2005	55	008	髙橋
ミウラ服装	NG2008	55	008	髙橋
ダンディ　マルマン	SI2011	55	008	髙橋
モードカジュアル　オギノ	ＳＩ2012	55	008	髙橋
フタバ	ＳＩ2013	55	008	髙橋
プロローグ	ＳＩ2015	60	008	髙橋
ラフェスタ	ＳＩ2016	55	008	髙橋
エディトリアル	SI2018	60	008	髙橋
シラクラ	TK2001	55	008	髙橋
モンアミ	TK2002	55	008	髙橋
タカヤマ	TK2024	55	008	髙橋
ココロショップ	TK2026	60	008	髙橋
フクダ洋服店	TK2038	55	008	髙橋
双葉屋洋服店	TK2039	55	008	髙橋
メゾン・エ・ココ	TK2042	55	008	髙橋
和田屋	TK2050	55	008	髙橋
モードギャラリー　やまもと	TK2055	55	008	髙橋"""

def main():
    count = 0
    with db_conn() as conn, conn.cursor() as cur:
        for line in raw_data.strip().split("\n"):
            parts = line.split("\t")
            if len(parts) >= 3:
                name = parts[0].strip()
                code = parts[1].strip()
                try:
                    rate = int(parts[2].strip())
                except ValueError:
                    rate = 35

                if not name:
                    continue

                try:
                    cur.execute("""
                        INSERT INTO customers (name, code, discount_rate)
                        VALUES (%s, %s, %s)
                        ON CONFLICT (name) DO UPDATE SET
                            code = EXCLUDED.code,
                            discount_rate = EXCLUDED.discount_rate;
                    """, (name, code, rate))
                    count += 1
                except Exception as e:
                    print(f"Failed to insert {name}: {e}")
                    conn.rollback()
                    continue
        conn.commit()
    print(f"Successfully inserted/updated {count} customers.")

if __name__ == "__main__":
    main()
