import marimo

__generated_with = "0.23.0"
app = marimo.App(width="medium")


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    mo.md(データを全区分、区分ごとに選択できるようにしたバージョン)
    mo.md(メモ：データをクレンジングする必要がある。５．わからない。というデータを除くとどうなるか？)
    """)
    return


@app.cell
def _():
    import marimo as mo
    import os
    import urllib.request
    import pandas as pd
    import matplotlib.pyplot as plt
    import seaborn as sns
    import pandas as pd
    from itertools import combinations
    import pingouin as pg
    import matplotlib.font_manager as fm
    import japanize_matplotlib
    import io

    ###フォント一覧を表示###

    # システムにインストールされているフォント名の一覧を取得
    fonts = fm.findSystemFonts()

    # フォント名（Family Name）をリストで表示
    font_names = [fm.FontProperties(fname=f).get_name() for f in fonts]

    # 重複を除いてアルファベット順に表示
    for name in sorted(set(font_names)):
        print(f"\"{name}\"")

    plt.rcParams['font.family'] = 'Noto Sans CJK JP' # インストールしたフォント名
    plt.rcParams['font.family'] = 'IPAexGothic'

    plt.title("日本語テスト")
    plt.show()
    ####


    # ファイルのダウンロード
    url = "https://www.ipa.go.jp/digital/software-survey/software-engineering/j5u9nn000000hkhs-att/software2025-result-data-id.csv"
    filename = "software2025-result-data-id.csv"

    # セル2: ダウンロード処理
    if not os.path.exists(filename):
        urllib.request.urlretrieve(url, filename)
        status = f"✅ 新しくダウンロードしました: {filename}"
    else:
        status = f"ℹ️ すでに存在するためスキップしました: {filename}"

    mo.md(status)
    return combinations, filename, mo, os, pd, pg, plt, sns, urllib


@app.cell
def _(os, pd, urllib):
    #設問リストのダウンロード
    def download_question_list():
        # 1. ExcelファイルのURLを指定
        url = "https://www.ipa.go.jp/digital/software-survey/software-engineering/j5u9nn000000hkhs-att/software2025-c-questions.xlsx"
        file_path = "software2025-c-questions.xlsx" # 保存するファイル名

        # 1. ファイルが存在するかチェック
        if not os.path.exists(file_path):
            print("ファイルをダウンロード中...")
            with urllib.request.urlopen(url) as response:
                content = response.read()
            # ローカルに保存
            with open(file_path, "wb") as f:
                f.write(content)
        else:
            print("既存のファイルを使用します。")

        # 2. 保存した（または既存の）ファイルから読み込み
        print("データを読み込みます")
        df_questions = pd.read_excel(file_path, sheet_name="設問一覧")
        return df_questions

    df_questions = download_question_list()
    return (df_questions,)


@app.cell
def _(filename, mo, pd):
    try:
        mo.md("欠損値を持つレコードをレコードから除去する。使い方には注意する！")
        df = pd.read_csv(filename)
        #これをやると、有効なレコードが0件になってしまう。
        #df = pd.read_csv(filename).dropna()

        # ○ 今回使う2つのカラムに絞って、両方埋まっている行だけ残す
        #target_cols = ['column_x', 'column_y']
        #df_clean = df.dropna(subset=target_cols)
        #という感じにする必要がある。
        # 最初の5行を表示して確認
        #print(df.head())
    except Exception as e:
        print(f"読み込みエラー: {e}")

    print(f"読み込み成功！ shape={df.shape}")
    return (df,)


@app.cell
def _(df):
    for index, row in df.head(5).iterrows():
        print(f"行番号: {index}, Q1-8.売上規模(1.売上高　（単体）): {row['Q1-8.売上規模(1.売上高　（単体）)']}, Q2-1.DXの取組状況: {row['Q2-1.DXの取組状況']}") # 列名を指定
    return


@app.cell
def _():
    """
    # Google ColabやLinux環境でよく使われる日本語フォントを指定
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Noto Sans CJK JP', 'IPAexGothic', 'Hiragino Sans', 'Yu Gothic', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False # 負の符号の文字化け防止

    x='Q1-8.売上規模(1.売上高　（単体）)'
    y='Q2-1.DXの取組状況'
    # 1. 相関係数の計算（デフォルトはピアソンの相関係数）
    #correlation = df['Q1-8.売上規模(1.売上高　（単体）)'].corr(df['Q2-1.DXの取組状況'])
    #method='spearman'
    #correlation = df['Q1-8.売上規模(1.売上高　（単体）)'].corr(df['Q2-1.DXの取組状況'],method='spearman')
    #method='kendall'
    #                    x軸                                    y軸
    correlation = df['Q1-8.売上規模(1.売上高　（単体）)'].corr(df['Q2-1.DXの取組状況'],method='kendall')
    ct = pd.crosstab(df[x], df[y])

    print(f"相関係数: {correlation:.4f}")

    # 2. 散布図の作成
    #plt.figure(figsize=(8, 6))
    #sns.scatterplot(data=df, x='Q1-8.売上規模(1.売上高　（単体）)', y='Q2-1.DXの取組状況')

    #sns.stripplot(data=df, x=x, y=y, jitter=0.25, alpha=0.5)
    sns.heatmap(ct, annot=True, fmt='d', cmap='Blues')
    plt.gca().invert_yaxis()

    # グラフのタイトルやラベル
    plt.title(f"Scatter Plot (Correlation: {correlation:.2f})")
    plt.xlabel("Column Q1-8.")
    plt.ylabel("Column Q2-1.")
    plt.grid(True)
    plt.show()
    """
    return


@app.cell
def _():
    #全組み合わせをグラフ描画しようとしたらNG。100万個以上になる！のでやめとく。
    """
    # 1. 除外したいラベルをリストで指定（完全一致）
    exclude_cols = ['ID', 'タイムスタンプ', '備考']

    # 2. 自動で全カラムを取得し、除外リストに含まれないものだけ残す
    # ついでに「Unnamed」を含む列なども自動で弾く設定にしています
    target_cols = [
        col for col in df.columns 
        if col not in exclude_cols and "Unnamed" not in col
    ]

    # 3. 全ての組み合わせを作成
    pairs = list(combinations(target_cols, 2))
    n_pairs = len(pairs)

    # 4. 描画設定（横に3つ並べる）
    ncols = 3
    nrows = (n_pairs + ncols - 1) // ncols
    fig, axes = plt.subplots(nrows=nrows, ncols=ncols, figsize=(6 * ncols, 5 * nrows))

    # グラフが1つしかない場合でも iterable にするために flatten
    if n_pairs == 1:
        axes = [axes]
    else:
        axes = axes.flatten()

    # 5. ループで描画
    for i, (col_x, col_y) in enumerate(pairs):
        ct = pd.crosstab(df[col_y], df[col_x])
        sns.heatmap(ct, annot=True, fmt='d', cmap='YlGnBu', ax=axes[i], cbar=False)

        axes[i].set_title(f"X: {col_x}\nY: {col_y}", fontsize=9)
        axes[i].invert_yaxis()

    # 余った枠を非表示
    for j in range(i + 1, len(axes)):
        axes[j].set_visible(False)

    plt.tight_layout()
    plt.show()
    """
    return


@app.cell
def _(pd, plt, sns):
    #ヒートマップグラフ定義
    def show_graph(df, col_x, col_y):
        ct = pd.crosstab(df[col_x], df[col_y])
    # 2. 散布図の作成
    #plt.figure(figsize=(8, 6))
    #sns.scatterplot(data=df, x='Q1-8.売上規模(1.売上高　（単体）)', y='Q2-1.DXの取組状況')

    #sns.stripplot(data=df, x=x, y=y, jitter=0.25, alpha=0.5)
        sns.heatmap(ct, annot=True, fmt='d', cmap='Blues')
        plt.gca().invert_yaxis()

        # グラフのタイトルやラベル
        plt.title(f"map")
        plt.xlabel(f"Column {col_x}")
        plt.ylabel(f"Column {col_y}")
        plt.grid(True)
        plt.show()

    return


@app.cell
def _(mo, pg):
    def find_spurious_correlation(selected_row_df, original_df, all_cols):
        if selected_row_df.empty:
            return mo.md("テーブルから行を選択してください")

        row = selected_row_df.iloc[0]
        col_x, col_y, orig_corr = row['column_x'], row['column_y'], row['corr']
        findings = []
        print(col_x,col_y)

        for potential_z in all_cols:
            if potential_z in [col_x, col_y]: continue

            # 3つの列を抜き出して欠損値を除去
            temp_df = original_df[[col_x, col_y, potential_z]].dropna()

            # 数値データでない場合は一旦スキップ（あるいは数値化が必要）
            # また、すべての値が同じ（分散ゼロ）列も計算できないため除外
            if temp_df[potential_z].nunique() <= 1:
                continue

            try:
                # 偏相関の計算
                res = pg.partial_corr(data=temp_df, x=col_x, y=col_y, covar=potential_z, method='kendall')
                new_corr = res['r'].values[0]

                # 元の相関の半分以下になったら「黒幕」候補
                if abs(new_corr) < abs(orig_corr) * 0.5:
                    findings.append(f"🔍 **{potential_z}** を考慮すると相関が低下 ({orig_corr:.3f} → **{new_corr:.3f}**)")
            except Exception:
                # SVDエラーなどの数学的エラーは無視して次へ
                continue

        if not findings:
            return mo.md("明確な「第3の変数（黒幕）」は見つかりませんでした。")

        return mo.vstack([mo.md("### 偽相関の可能性がある要因"), *[mo.md(f) for f in findings]])


    return (find_spurious_correlation,)


@app.cell
def _(mo, pd):
    # サンプルデータ
    sangyou_kubun_mapping = {
        1: "農業・林業",
        2: "漁業",
        3: "鉱業、採石業、砂利採取業",
        4: "建設業",
        5: "製造業",
        6: "電気・ガス・熱供給・水道業",
        7: "情報通信業",
        8: "運輸業・郵便業",
        9: "卸売業、小売業",
        10: "金融業、保険業",
        11: "不動産業、物品賃貸業",
        12: "宿泊業、飲食サービス業",
        13: "生活関連サービス業、娯楽業",
        14: "教育、学習支援業",
        15: "医療、福祉",
        16: "複合サービス事業",
        17: "サービス業（他に分類されないもの）",
        18: "公務（他に分類されるものを除く）",
        19: "分類不能の産業"
    }

    df_sangyou_kubun_mapping = pd.DataFrame(
        list(sangyou_kubun_mapping.items()), 
        columns=["産業区分コード", "産業区分"]
    )

    # 1. 選択肢に「全業種」を追加
    options = ["全区分"] + df_sangyou_kubun_mapping["産業区分"].tolist()

    category = mo.ui.dropdown(
        options=options, 
        value="全区分",  # 初期値を全業種に設定
        label="産業区分を選択してください"
    )
    category
    return category, sangyou_kubun_mapping


@app.cell
def _(
    category,
    combinations,
    df,
    mo,
    pd,
    sangyou_kubun,
    sangyou_kubun_mapping,
):
    #そこで、まずは、全組み合わせで超絶力技で相関係数だけを計算する

    # 選択された産業区分に応じてテーブルをフィルタリング
    if category.value == "全区分":
        df_temp = df#全データ
    else:
        for i in sangyou_kubun:
            selected_key = [k for k, v in sangyou_kubun_mapping.items() if v == category.value]
            selected_key = selected_key[0] if selected_key else None
            break
        #dfに入っているQ1-3.産業区分はコード。これを一旦文字列に変換して、ユーザ選択と比較
        df_temp = df[df["Q1-3.産業区分"] == selected_key]


    results = []


    # 1. 除外したいラベルをリストで指定（完全一致）
    exclude_cols = ['Q1-3.産業区分', 'Q1-4.企業種別', 'Q1-5.所在地', 'Q1-6.設立年数']
    #exclude_cols = []

    # 2. 自動で全カラムを取得し、除外リストに含まれないものだけ残す
    # ついでに「Unnamed」を含む列なども自動で弾く設定にしています
    target_cols = [
        col for col in df_temp.columns 
        if col not in exclude_cols and "Unnamed" not in col
    ]

    # 3. 全ての組み合わせを作成
    pairs = list(combinations(target_cols, 2))
    n_pairs = len(pairs)

    # 5. ループで描画
    for i, (col_x, col_y) in enumerate(pairs):
        # 1. 相関係数の計算（デフォルトはピアソンの相関係数）
        #correlation = df_temp['Q1-8.売上規模(1.売上高　（単体）)'].corr(df_temp['Q2-1.DXの取組状況'])
        #method='spearman'
        #correlation = df_temp['Q1-8.売上規模(1.売上高　（単体）)'].corr(df_temp['Q2-1.DXの取組状況'],method='spearman')
        #method='kendall'
        #                    x軸                                    y軸
        correlation = df_temp[col_x].corr(df_temp[col_y],method='kendall')    
        #print(f"相関係数: {correlation:.4f}")

        # 辞書形式で結果を保存
        results.append({
            'column_x': col_x,
            'column_y': col_y,
            'corr': correlation
        })

    # リストをDataFrameに変換
    corr_df = pd.DataFrame(results)
    corr_df = corr_df.dropna(subset=['corr'])

    # 相関係数の降順（高い順）に並び替え
    # 絶対値で評価したい場合は .abs() を使うこともありますが、まずはそのままの数値で。
    corr_df_sorted = corr_df.sort_values(by='corr', ascending=False)

    # 結果を表示
    #print(corr_df_sorted)
    table = mo.ui.table(
        corr_df_sorted, 
        page_size=100,
        label=f"相関係数ランキング（総当たり）産業区分={category.value}"
    )
    table
    return df_temp, table, target_cols


@app.cell
def _(df_temp, find_spurious_correlation, table, target_cols):
    find_spurious_correlation(table.value, df_temp, target_cols)
    #これやると、相関係数行列はfinite valuesのみだって怒られる。
    #sns.clustermap(df.corr())
    return


@app.cell
def _(df_temp, mo, render_selected_heatmap, table):
    # 関数を呼び出して表示
    # t.value はテーブルで選択されたDataFrameが入っている想定です
    hmap_fig, (exp_x, exp_y) = render_selected_heatmap(table.value, df_temp)
    text_x = exp_x.iloc[0] if not exp_x.empty else "情報なし"
    text_y = exp_y.iloc[0] if not exp_y.empty else "情報なし"
    # 1. 2つのテキストを「縦（Vertical）」に並べる
    # フォントサイズを小さくするために Markdown の <sub> や 小さめの見出しを使います
    text_content = mo.vstack([
        mo.md(f"**【X軸】**<br><small>{text_x}</small>"),
        mo.md(f"**【Y軸】**<br><small>{text_y}</small>")
    ])

    # 2. 横並びにする際、グラフとテキストの幅の比率を指定する
    # widths=[2, 1] とすると、グラフに2/3、テキストに1/3の幅を割り当てます
    layout = mo.hstack(
        [hmap_fig, text_content], 
        widths=[2, 1],  # グラフを大きく、テキストを狭く
        align="start"
    )

    layout
    return


@app.cell
def _(df, plt, sangyou_kubun_mapping):
    # カラム「産業区分」のヒストグラムを作成
    #df['Q1-3.産業区分'].value_counts().sort_index().plot.bar()
    #df['Q1-3.産業区分'].value_counts().plot.bar()
    #df['Q1-3.産業区分'].map(mapping).value_counts().plot.bar()

    # グラフを表示するための設定
    #plt.title("産業区分別の件数（降順）")
    #plt.show()
    # フォント設定
    #plt.rcParams['font.family'] = 'IPAexGothic'

    # 1. データの集計
    counts = df['Q1-3.産業区分'].map(sangyou_kubun_mapping).value_counts()

    # 2. グラフの描画（横棒グラフ）
    ax = counts.plot.barh()
    ax.invert_yaxis()  # 上から多い順にする

    # ★ 3. 棒の先端に数値を表示する
    ax.bar_label(ax.containers[0], padding=3)

    # 見栄えの調整
    plt.title("産業区分別の分布（件数表示）")
    plt.xlabel("件数")
    plt.tight_layout()  # ラベルが画面外にはみ出るのを防ぐ

    plt.show()
    return


@app.cell
def _(combinations, pd, render_heatmap):
    def targeted_analy1_aux(all_df,target_sanngyou_kubun_code,target_cols):
        results = []
        figs = []

        #target_cols = ["Q2-1.DXの取組状況","Q2-2.DXの成果","Q2-6.DXを推進する人材の状況","Q5-2.レガシーシステムの影響"]
        #gyousyu = ["製造業","建設業","卸売業、小売業","サービス業（他に分類されないもの）"]

        df = all_df[all_df['Q1-3.産業区分'] == target_sanngyou_kubun_code]
        pairs = list(combinations(target_cols, 2))
        # 5. ループで描画
        for i, (col_x, col_y) in enumerate(pairs):
            #                    x軸                                    y軸
            correlation = df[col_x].corr(df[col_y],method='kendall')    
            #print(f"相関係数: {correlation:.4f}")

            # 辞書形式で結果を保存
            results.append({
                'column_x': col_x,
                'column_y': col_y,
                'corr': correlation
            })

        # リストをDataFrameに変換
        corr_df = pd.DataFrame(results)

        # 相関係数の降順（高い順）に並び替え
        # 絶対値で評価したい場合は .abs() を使うこともありますが、まずはそのままの数値で。
        corr_df_sorted = corr_df.sort_values(by='corr', ascending=False)

        for selected in corr_df_sorted.iterrows():
            idx, row = selected
            fig = render_heatmap(row["column_x"], row["column_y"], row["corr"], df)
            figs.append(fig)

        return figs

    return


@app.cell
def _(combinations, pd, render_heatmap):

    def default_find_pair(target_cols, all_df):
        return list(combinations(target_cols, 2))

    def find_pair_other_all_cols(target_cols, all_df):
        output = []
        # all_df.columns から target_cols に含まれないものだけを抽出
        remaining_cols = [col for col in all_df.columns if col not in target_cols]
        for x in target_cols:
            for y in remaining_cols:
                output.append((x,y))
        return output


    #targeted_analy1_auxのpairsを関数ポインタにしたもの
    def targeted_analy1_aux2(all_df,target_sanngyou_kubun_code,target_cols,find_pair_func=None):
        results = []
        figs = []

        if find_pair_func is None:
            find_pair_func = default_find_pair

        #target_cols = ["Q2-1.DXの取組状況","Q2-2.DXの成果","Q2-6.DXを推進する人材の状況","Q5-2.レガシーシステムの影響"]
        #gyousyu = ["製造業","建設業","卸売業、小売業","サービス業（他に分類されないもの）"]

        df = all_df[all_df['Q1-3.産業区分'] == target_sanngyou_kubun_code]
        pairs = find_pair_func(target_cols, all_df)
        #pairs = list(combinations(target_cols, 2))
        # 5. ループで描画
        for i, (col_x, col_y) in enumerate(pairs):
            #                    x軸                                    y軸
            correlation = df[col_x].corr(df[col_y],method='kendall')    
            #print(f"相関係数: {correlation:.4f}")

            # 辞書形式で結果を保存
            results.append({
                'column_x': col_x,
                'column_y': col_y,
                'corr': correlation
            })

        # リストをDataFrameに変換
        corr_df = pd.DataFrame(results)

        # 相関係数の降順（高い順）に並び替え
        # 絶対値で評価したい場合は .abs() を使うこともありますが、まずはそのままの数値で。
        corr_df_sorted = corr_df.sort_values(by='corr', ascending=False)

        for selected in corr_df_sorted.iterrows():
            idx, row = selected
            if pd.isna(row["corr"]):
                continue
            fig = render_heatmap(row["column_x"], row["column_y"], row["corr"], df)
            figs.append(fig)

        return figs

    return (targeted_analy1_aux2,)


@app.cell
def _(df, mo, sangyou_kubun_mapping, targeted_analy1_aux2):
    def targeted_analy1(sangyou_kubun, sangyou_kubun_mapping, find_pair_func = None):
        figs = []
        target_cols = ["Q2-1.DXの取組状況","Q2-2.DXの成果","Q2-6.DXを推進する人材の状況","Q5-2.レガシーシステムの影響"]

        sangyou_kubun_codes = []
        for i in sangyou_kubun:
            key = next((k for k, v in sangyou_kubun_mapping.items() if v == i)), 
            sangyou_kubun_codes.append(key)


        for i in sangyou_kubun_codes:
            #産業区分コードごとにグラフの配列を分ける
            figs.append(targeted_analy1_aux2(df, i, target_cols,find_pair_func=find_pair_func))

        # 業界名のリスト　＝　sangyou_kubun
        # figs = [[fig, fig], [fig, fig], ...] (お手元のデータ)
        figs_list = figs
        tabs_dict = {}
        for name1, figs in zip(sangyou_kubun, figs_list):
            # 各業界の中にある複数のfigをvstackで縦に並べる
            tabs_dict[name1] = mo.vstack(figs)

        return tabs_dict


    sangyou_kubun = ["製造業","建設業","卸売業、小売業","サービス業（他に分類されないもの）","金融業、保険業","運輸業・郵便業"]
    tabs_dict = targeted_analy1(sangyou_kubun, sangyou_kubun_mapping)
    return sangyou_kubun, tabs_dict


@app.cell
def _(mo, tabs_dict):
    # セルの1番下にこれだけ書く
    # タブとして表示
    mo.tabs(tabs_dict)
    return


@app.cell
def _():
    #tabs_dict2 = targeted_analy1(sangyou_kubun, sangyou_kubun_mapping,find_pair_func=find_pair_other_all_cols)
    #mo.tabs(tabs_dict2)
    return


@app.cell
def _(df_questions, mo, pd, plt, sns):

    def select_questions_explanation(column_x, column_y):
        target_q = column_x.split(".")[0]
        #filtered_df = df_questions[df_questions.iloc[:, 0].str.startswith(target_q, na=False)]
        filtered_df = df_questions[df_questions.iloc[:, 0] == target_q]
        exp_x = filtered_df["選択肢"]

        target_q = column_y.split(".")[0]
        #filtered_df = df_questions[df_questions.iloc[:, 0].str.startswith(target_q, na=False)]
        filtered_df = df_questions[df_questions.iloc[:, 0] == target_q]
        exp_y = filtered_df["選択肢"]

        """
        rosstab は 第1引数が「行（縦軸・Y）」、第2引数が「列（横軸・X）」 になります。一方、その下の描画や説明の取得では c_x が X軸（横）として扱われているため、もし crosstab の中で c_x を第2引数（列）に渡していると、ヒートマップの見た目上の横軸は c_x になりますが、データの並びや解釈が混乱しやすくなります。
        """
        return (exp_x, exp_y)
        #return (exp_y, exp_x) #ってことで、ひっくりかえしてやる

    def render_selected_heatmap(selected_df, source_df):
        """
        選択された1行のデータからヒートマップを生成する関数
        """
        # 選択がない場合はメッセージを返す
        if selected_df.empty:
            return (mo.md("💡 **テーブルから行を選択してください。詳細なヒートマップが表示されます。**"), (pd.Series(dtype='str'),pd.Series(dtype='str')))

        # 1行目を取り出す
        row = selected_df.iloc[0]
        c_x = row['column_x']
        c_y = row['column_y']
        c_val = row['corr']

        # 描画処理（変数名は関数内のみで有効）
        fig, ax = plt.subplots(figsize=(7, 5))
        ct = pd.crosstab(source_df[c_y], source_df[c_x])
        sns.heatmap(ct, annot=True, fmt='d', cmap='YlGnBu', ax=ax, cbar=False)

        ax.set_title(f"詳細分析: {c_x} vs {c_y}\n(τ: {c_val:.4f})", fontsize=10)
        ax.invert_yaxis()

        # marimo表示用に変換
        output = mo.as_html(fig)
        plt.close(fig)
        return (output, select_questions_explanation(c_x, c_y))

    def render_heatmap(column_x, column_y, corr, source_df):
        c_x = column_x
        c_y = column_y
        c_val = corr

        # 描画処理（変数名は関数内のみで有効）
        fig, ax = plt.subplots(figsize=(7, 5))
        ct = pd.crosstab(source_df[c_y], source_df[c_x])
        sns.heatmap(ct, annot=True, fmt='d', cmap='YlGnBu', ax=ax, cbar=False)

        ax.set_title(f"詳細分析: {c_x} vs {c_y}\n(τ: {c_val:.4f})", fontsize=10)
        ax.invert_yaxis()

        # marimo表示用に変換
        output = mo.as_html(fig)
        #plt
        plt.close(fig)
        #output = fig
        return output   


    return render_heatmap, render_selected_heatmap


if __name__ == "__main__":
    app.run()
