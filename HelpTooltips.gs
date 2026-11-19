/**
 * ❔ HelpTooltips.gs — Tooltip Logic Definitions
 * ─────────────────────────────────────────────
 * 📋 Purpose:
 * Houses all tooltip function definitions for the “Tooltips & Help” menu in the 
 * Japanese Study Tracker system. Each function provides contextual guidance via 
 * alert boxes or sidebars, triggered by corresponding menu items.
 *
 * 🛠️ Key Features:
 * - Clear, themed tooltip content for each major workflow or system component
 * - Bilingual explanations (English + polite Japanese) for learning reinforcement
 * - Lightweight: no dependencies beyond standard Spreadsheet UI alerts
 *
 * 🧩 Typical Functions:
 * - `showTaskSyncTooltip()`
 * - `showLoggerTooltip()`
 * - `showDashboardOverviewTooltip()`
 * - … and many more
 *
 * 📌 Dependencies:
 * The menu structure that calls these functions is defined in `HelpTooltipsMenu.gs`
 *
 * 🛠 Maintainer Notes:
 * Add new tooltip functions here or segment them by theme into sub-files 
 * (e.g., `Tooltips_Syncing.gs`, `Tooltips_Metrics.gs`) for long-term scalability.
 */

// ─────────────────────────────────────────────


// ✅ Logger System Explanation
function showLoggerTooltip() {
  const msg =
    "🧾 STUDY LOGGERS\n\n" +
    "- Use these to manually log any study activity.\n" +
    "- All entries go into the 'Learning Log' sheet — your core record of progress.\n" +
    "- This log is permanent and resilient — safe from calendar or task deletions.\n" +
    "- Log every meaningful session to build long-term habits and data accuracy.\n\n" +
    "✅ Pro Tip: Log Kanji, vocab, grammar, structured lessons, or even fun immersion.\n\n" +
    "——————————————\n\n" +
    "🧾 スタディ・ロガー（学習記録）\n\n" +
    "・このツールで、学習活動を手動で記録できます。\n" +
    "・すべての記録は「Learning Log（学習ログ）」シートに保存されます。\n" +
    "・このログは恒久的で信頼性があり、タスクやカレンダーの削除からも保護されます。\n" +
    "・意味のある学習セッションは必ず記録して、習慣とデータ精度を築きましょう。\n\n" +
    "✅ プロのヒント：漢字、語彙、文法、レッスン、楽しい没入型の学習も記録できます。\n";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Date Assignment Tooltip
function showDateAssignTooltip() {
  const msg =
    "📅 ASSIGN DATES TO PHASE\n\n" +
    "- Fills out the selected Phase block in your Master Lesson Plan (columns A–C).\n" +
    "- Pulls the start date from Dashboard B4 and applies 90 sequential days to the phase in Dashboard B3.\n" +
    "- Required BEFORE syncing tasks or calendar events — this maps each lesson to a real-world date.\n\n" +
    "⚠️ Check that Dashboard cell B3 (Current Phase) and B4 (Start Date) are both set correctly.\n\n" +
    "📅 フェーズに日付を割り当て\n\n" +
    "- 『Master Lesson Plan』シートの選択されたフェーズブロック（列A〜C）に日付を割り当てます。\n" +
    "- ダッシュボードのB4から開始日を取得し、B3で選択されたフェーズに90日分の連続日付を適用します。\n" +
    "- タスクやカレンダーを同期する前に、この処理を必ず実行してください。これにより、各レッスンが実際の日付に対応します。\n\n" +
    "⚠️ ダッシュボードのセルB3（現在のフェーズ）とB4（開始日）が正しく設定されていることを確認してください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Task Sync Tooltip
function showTaskSyncTooltip() {
  const msg =
    "📋 SYNC TASKS (1–180)\n\n" +
    "- Syncs structured daily tasks from the Master Lesson Plan sheet to your Google Tasks list for the selected phase.\n" +
    "- Tasks are chunked into 30-day blocks to prevent exceeding Google API rate limits.\n" +
    "- Make sure you’ve assigned dates to the phase (via Dashboard) and created the necessary task lists first.\n\n" +
    "🛑 Previously synced tasks are NOT automatically removed — use 'Wipe All Task Lists' before re-syncing.\n\n" +
    "📋 タスクを同期（1～180日）\n\n" +
    "- 『Master Lesson Plan』シートから、選択されたフェーズの構造化された日次タスクを Google Tasks に同期します。\n" +
    "- タスクは、Google API の制限を避けるため、30日ごとのチャンクに分けて同期されます。\n" +
    "- 事前にダッシュボードでフェーズの日付を割り当て、必要なタスクリストを作成しておいてください。\n\n" +
    "🛑 既存のタスクは自動的に削除されません。再同期する前に「🧨 Wipe All Task Lists」でタスクを削除してください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Calendar Sync Tooltip
function showCalendarSyncTooltip() {
  const msg =
    "📅 SYNC CALENDAR EVENTS (1–180)\n\n" +
    "- Syncs structured study sessions from the Master Lesson Plan to your Google Calendar as visible, time-blocked events.\n" +
    "- Uses 30-day chunks to respect Google Calendar API limits.\n" +
    "- Assign dates to your phase via the Dashboard BEFORE running any calendar syncs.\n\n" +
    "🧽 Re-syncing? Run 'Clean Calendar Events' first to avoid duplicates or clutter.\n\n" +
    "📅 カレンダーイベントの同期（1～180日）\n\n" +
    "- 『Master Lesson Plan』シートの学習セッションを、時間ブロック付きのイベントとして Google カレンダーに同期します。\n" +
    "- Google カレンダー API の制限を考慮し、30日ごとのチャンクで処理されます。\n" +
    "- カレンダーを同期する前に、必ずダッシュボードでフェーズの日付を割り当ててください。\n\n" +
    "🧽 再同期する場合は、重複やごちゃごちゃを避けるために「🗑 Clean Calendar Events」を先に実行してください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Cleanup Tooltip
function showCleanupTooltip() {
  const msg =
    "🧼 CLEAN-UP & RESYNC\n\n" +
    "- Wipe All Task Lists → Deletes all synced Google Task LISTS at the root level, which removes all associated tasks.\n" +
    "- Clean Calendar Events → Removes all study-related events from your Google Calendar.\n\n" +
    "📌 After wiping tasks:\n" +
    "• Re-run 'Create Task Lists' to re-establish your phase task containers.\n" +
    "• Then re-sync tasks and calendar events as needed.\n\n" +
    "⚠️ Why this is required:\n" +
    "• Google’s APIs heavily cache task list states and prevent clean overwriting.\n" +
    "• Attempts to overwrite tasks directly caused instability and duplication.\n" +
    "• A 'one-button full reset' was too slow due to forced Google wait periods — this manual step is faster and safer.\n\n" +
    "✅ These actions are non-destructive to your data — your Learning Log and curriculum stay fully intact.\n" +
    "🧠 Completed task history is not preserved, so check off items again manually if needed.\n\n" +
    "🧼 クリーンアップと再同期\n\n" +
    "- 「Wipe All Task Lists」→ Google タスクのリストごと削除（中のタスクもすべて消去されます）\n" +
    "- 「Clean Calendar Events」→ Google カレンダーに同期された学習イベントをすべて削除します\n\n" +
    "📌 タスク削除後の手順：\n" +
    "・『Create Task Lists』を再実行して、各フェーズのタスクリストを作り直します\n" +
    "・その後、タスクとカレンダーを再同期します\n\n" +
    "⚠️ なぜこの方法なのか：\n" +
    "・Google の API はタスクリスト状態をキャッシュしており、上書きや再構築に不向き\n" +
    "・強引な上書きはタスクの重複や同期失敗の原因になりました\n" +
    "・一括処理（ワンボタン方式）は Google による待機時間が発生し、逆に時間がかかります\n\n" +
    "✅ この処理は安全です。学習ログやカリキュラムの内容は一切変更されません。\n" +
    "🧠 削除後は、完了済みのタスクは手動でチェックし直してください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ System-Wide Tips & Guardrails
function showSystemTipsTooltip() {
  const msg =
    "💡 SYSTEM TIPS & WARNINGS\n\n" +
    "- The Master Sheet is your source of truth. DO NOT manually edit tasks or calendar events — always use the sync tools.\n" +
    "- If you manually change or delete items in Google Tasks or Calendar, they will fall out of sync with the system.\n" +
    "- Follow the correct sync order every time:\n" +
    "   1. Assign Dates\n" +
    "   2. Create Task Lists\n" +
    "   3. Sync Tasks\n" +
    "   4. Sync Calendar (optional)\n\n" +
    "✅ Log your study sessions consistently — your Learning Log is the permanent, reliable record.\n\n" +
    "💡 システムに関するヒントと注意事項\n\n" +
    "- 「Master Sheet（マスターシート）」は信頼できるデータの基準です。タスクやカレンダーイベントを手動で編集せず、必ず同期ツールを使ってください。\n" +
    "- Google Tasks や Google カレンダーで手動で変更・削除を行うと、システムとの同期が崩れます。\n" +
    "- 毎回、以下の順番で操作してください：\n" +
    "   1. Assign Dates（日付の割り当て）\n" +
    "   2. Create Task Lists（タスクリストの作成）\n" +
    "   3. Sync Tasks（タスクの同期）\n" +
    "   4. Sync Calendar（カレンダーの同期：任意）\n\n" +
    "✅ 学習セッションを定期的に記録しましょう。「Learning Log（学習ログ）」は信頼できる履歴として機能します。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ UTC Date Warning Tooltip
function showDateWarningTooltip() {
  const msg =
    "⚠️ TIMEZONE NOTICE: Google uses UTC\n\n" +
    "- Google Apps Script stores dates in **UTC (Coordinated Universal Time)**.\n" +
    "- If you're in a timezone **behind UTC** (like most US timezones), logging today's date may display as the previous day.\n\n" +
    "➡️ To ensure correct display:\n" +
    "• Select the date one day ahead when logging (e.g., choose April 30 to record April 29).\n" +
    "• Then run 🩺 'Heal Session Anchors' from the menu.\n\n" +
    "✅ Your dashboard will automatically reflect the correct date after healing.\n" +
    "\n\n" +
    "⚠️ タイムゾーンに関する注意事項（UTC仕様）\n\n" +
    "- Google Apps Script は日付を **UTC（協定世界時）** で保存します。\n" +
    "- 日本やアメリカなど、UTCより遅れているタイムゾーンでは、\n" +
    "　今日の日付を選んでも前日として表示されることがあります。\n\n" +
    "➡️ 表示を正しくするために：\n" +
    "• 日付を1日進めて選択してください（例：4月29日を記録したい場合は4月30日を選択）。\n" +
    "• その後、メニューから 🩺『Heal Session Anchors（セッションアンカー修復）』を実行してください。\n\n" +
    "✅ 修復後、ダッシュボードには正しい日付が自動で反映されます。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Chunk Sync Tooltip
function showChunkSyncTooltip() {
  const msg =
    "🔁 SYNC BY CHUNK\n\n" +
    "- Sync 30-day blocks of tasks or calendar events using a number-based prompt.\n" +
    "- Enter a number 1–6 to sync the corresponding 30-day window (e.g., 3 = Days 61–90).\n" +
    "- Safer for avoiding Google API rate limits.\n\n" +
    "✅ Pro Tip: Clean tasks or calendar events first if re-syncing。\n\n" +
    "🔁 チャンクごとの同期\n\n" +
    "- 数字入力によって、30日単位のタスクやカレンダーイベントを同期します。\n" +
    "- 1〜6 の番号を入力して、該当する30日間のウィンドウを同期します（例：3 = 61日目〜90日目）。\n" +
    "- Google APIの制限を避けるための安全な方法です。\n\n" +
    "✅ プロのヒント：再同期する場合は、まずタスクやカレンダーイベントをクリーンアップしましょう。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Sheet ID and Phase Settings Tooltip
function showSheetPhaseHelpTooltip() {
  const msg =
    "📄 SHEET ID & PHASE SETTINGS\n\n" +
    "Your Dashboard controls where tasks are logged and synced from:\n\n" +
    "🆔 B27 = Spreadsheet ID\n" +
    "   → This connects to your Phase Sheet (e.g. where the Learning Log lives).\n\n" +
    "📘 B3 = Current Phase\n" +
    "   → Tells the system which Phase block in the Master Lesson Plan to use.\n\n" +
    "🗓️ B28 = Calendar ID\n" +
    "   → Tells the system which Google Calendar to sync your events to.\n\n" +
    "⚠️ All 3 must be valid for the system to work properly. If any are missing or wrong, syncs will fail or go to the wrong place.\n\n" +
    "✅ You can paste alternate spreadsheet/calendar IDs here if you're managing multiple setups。\n\n" +
    "📄 シートIDとフェーズ設定\n\n" +
    "ダッシュボードは、タスクの記録や同期先を管理します：\n\n" +
    "🆔 B27 = スプレッドシートID\n" +
    "   → 対象のフェーズシート（Learning Logが存在する場所）に接続します。\n\n" +
    "📘 B3 = 現在のフェーズ\n" +
    "   → 『Master Lesson Plan』のどのフェーズブロックを使うかをシステムに指示します。\n\n" +
    "🗓️ B28 = カレンダーID\n" +
    "   → 同期対象となるGoogleカレンダーを指定します。\n\n" +
    "⚠️ この3つすべてが正しく設定されていないと、同期に失敗したり、誤った場所に送られたりします。\n\n" +
    "✅ 複数のセットアップを管理している場合は、別のスプレッドシートやカレンダーIDをここに貼り付けることも可能です。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ System Flow Tooltip
function showSystemInternalsTooltip() {
  const msg =
    "🧠 SYSTEM FLOW OVERVIEW\n\n" +
    "This is how your Japanese Study Tracker works:\n\n" +
    "1️⃣ Dashboard B3 → Current Phase\n" +
    "   Controls which Phase’s lessons and logs are active.\n\n" +
    "2️⃣ Dashboard B4 → Phase Start Date\n" +
    "   Maps Master Lesson Plan days to real calendar dates.\n\n" +
    "3️⃣ Dashboard B27 → Sheet ID\n" +
    "   Routes all logs to the correct Phase Log Sheet.\n\n" +
    "4️⃣ Dashboard B28 → Calendar ID (Optional)\n" +
    "   Sends study sessions to the right Google Calendar.\n\n" +
    "📘 The Dashboard is your control center.\n" +
    "📝 Logging happens via forms → goes to your active Phase Log.\n" +
    "📊 Metrics pull live from the log — this is your source of truth.\n\n" +
    "🧾 You can assign dates to enable optional syncing.\n" +
    "🛠️ Update Dashboard auto-creates new logs as needed.\n\n" +
    "✅ Use the tools you want: log-only, or full syncing.\n\n" +
    "————————————————————————\n\n" +

    "🧠 システムの流れ（概要）\n\n" +
    "Japanese Study Tracker の連携構造：\n\n" +
    "1️⃣ ダッシュボード B3 → 現在のフェーズ\n" +
    "   アクティブなフェーズのレッスンとログを制御します。\n\n" +
    "2️⃣ ダッシュボード B4 → フェーズの開始日\n" +
    "   『Master Lesson Plan』の「Day」を実際の日付にマッピングします。\n\n" +
    "3️⃣ ダッシュボード B27 → シートID\n" +
    "   学習ログの保存先となるフェーズログシートを指定します。\n\n" +
    "4️⃣ ダッシュボード B28 → カレンダーID（任意）\n" +
    "   Googleカレンダーへの学習イベントの送信先を指定します。\n\n" +
    "📘 ダッシュボードがすべての中枢です。\n" +
    "📝 フォームで記録 → アクティブなログシートに保存されます。\n" +
    "📊 メトリクスは常にログから取得され、これが真実の情報源です。\n\n" +
    "🧾 必要なら日付を割り当てて、同期機能も使えます。\n" +
    "🛠️ 『Update Dashboard』でログ作成などの処理も自動化されます。\n\n" +
    "✅ シンプルに記録だけでもOK、同期まで使ってもOK。";
    
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Metrics Overview Tooltip
function showMetricsOverviewTooltip() {
  const msg =
    "📊 HOW METRICS WORK\n\n" +
    "- Metrics pull from your Learning Log (Columns A–I).\n" +
    "- Days Studied and Streaks are tracked based on unique day numbers logged (Column E), regardless of 'Completed' status.\n" +
    "- Completion Rate compares logged days (E or F) with your Master Lesson Plan — the ✅ checkbox isn't required.\n" +
    "- Kanji & Vocab entries are counted separately to track depth of language acquisition.\n" +
    "- All dashboard metrics refresh when you run 'Update Dashboard'.\n\n" +
    "📊 メトリクスの仕組み\n\n" +
    "- メトリクスは「Learning Log（学習ログ）」シートのA〜I列からデータを取得します。\n" +
    "- 「学習日数」と「連続記録（ストリーク）」は、E列に記録されたユニークな日数に基づいて計算されます。\n" +
    "- 完了率はE列またはF列の記録日をMaster Lesson Planと照合して算出します。チェックボックス列Hは不要です。\n" +
    "- 漢字と語彙の記録は分けてカウントされ、言語習得の深さを分析します。\n" +
    "- 「🧠 Update Dashboard」を実行すると、すべてのメトリクスが更新されます。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Streak Overview Tooltip
function showStreakTooltip() {
  const msg =
    "📈 WHAT'S A STREAK?\n\n" +
    "- Your streak tracks how many days in a row you've studied.\n" +
    "- The system allows a 1-day buffer: if you haven't logged today yet, your streak still continues — assuming you’ll study later.\n" +
    "- You lose your streak only if you skip a full day with no activity.\n\n" +
    "🔥 = 10+ days\n🌊 = 20+ days\n🌸 = 30+ days\n\n" +
    "📈 連続記録（ストリーク）とは？\n\n" +
    "- ストリークとは、何日連続で学習を記録したかを追跡する機能です。\n" +
    "- システムは1日の猶予（バッファ）を設けています。今日まだ記録していなくても、後で学習すればストリークは継続されます。\n" +
    "- 完全に学習のない日が1日でもあると、ストリークはリセットされます。\n\n" +
    "🔥 = 10日以上\n🌊 = 20日以上\n🌸 = 30日以上";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Sparkline Overview Tooltip
function showSparklineTooltip() {
  const msg =
    "🔥 HOW SPARKLINE WORKS\n\n" +
    "- This shows your study rhythm over the last 14 days.\n" +
    "- █ = Studied that day\n░ = Skipped that day\n\n" +
    "📘 It's a quick visual habit radar — stay consistent!\n\n" +
    "🔥 スパークラインとは？\n\n" +
    "- 直近14日間の学習リズムを可視化したグラフです。\n" +
    "- █ = 学習した日\n░ = 学習しなかった日\n\n" +
    "📘 習慣の可視化に役立つ簡易レーダーです。学習を継続していきましょう！";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Kanji vs. Vocab Tooltip
function showKanjiVocabTipTooltip() {
  const msg =
    "🧾 KANJI VS VOCAB ENTRIES\n\n" +
    "- Kanji = individual characters or focused writing practice.\n" +
    "- Vocabulary = full word study or usage examples.\n\n" +
    "📘 Tracking these separately helps reinforce memory and progress.\n\n" +
    "🧾 漢字と語彙のエントリーについて\n\n" +
    "- 漢字 = 単体の文字や筆記練習を目的とした記録。\n" +
    "- 語彙 = 単語全体の学習や使用例の記録。\n\n" +
    "📘 別々に記録することで記憶の定着や進捗の確認がしやすくなります。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Difficulty Rating Tooltip
function showDifficultyRatingTipTooltip() {
  const msg =
    "🎚️ DIFFICULTY RATINGS (1–5)\n\n" +
    "- Use this to rate how hard a session felt.\n" +
    "- 1 = Super Easy\n3 = Moderate\n5 = Very Difficult\n\n" +
    "📘 Helps spot burnout patterns or hard days when reviewing logs.\n\n" +
    "🎚️ 難易度評価（1〜5）\n\n" +
    "- セッションがどのくらい難しかったかを評価しましょう。\n" +
    "- 1 = とても簡単\n3 = 普通\n5 = とても難しい\n\n" +
    "📘 ログを見返すときに、疲労の傾向や学習の波を把握できます。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Dashboard Overview Tooltip
function showDashboardOverviewTooltip() {
  const msg =
    "📊 DASHBOARD OVERVIEW\n\n" +
    "- The Dashboard is your study control panel.\n" +
    "- It summarizes your streaks, progress, difficulty, study time, and more.\n" +
    "- All metrics update when you run '🧠 Update Dashboard'.\n\n" +
    "📌 Phase Selector (B3): Make sure this is set correctly — it controls which Phase is tracked.\n\n" +
    "📊 ダッシュボード概要\n\n" +
    "- ダッシュボードは、あなたの学習のコントロールパネルです。\n" +
    "- ストリーク、進捗、難易度、学習時間などを集約して表示します。\n" +
    "- 「🧠 Update Dashboard」を実行すると、すべてのメトリクスが最新になります。\n\n" +
    "📌 セルB3の「Phase Selector」で追跡するフェーズが決まります。正しく選択されていることを確認してください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Dashboard Update Tooltip
function showDashboardRefreshTooltip() {
  const msg =
    "🔄 WHEN TO UPDATE THE DASHBOARD\n\n" +
    "- Run '🧠 Update Dashboard' after any of these:\n" +
    "  • Logging new study sessions\n" +
    "  • Healing session anchors\n" +
    "  • Syncing tasks or calendar events\n\n" +
    "📌 This ensures your streaks and metrics reflect your latest progress.\n\n" +
    "🔄 ダッシュボードの更新タイミング\n\n" +
    "- 以下の操作後には、必ず「🧠 Update Dashboard」を実行してください：\n" +
    "  ・新しい学習セッションの記録後\n" +
    "  ・セッションアンカーの修復後\n" +
    "  ・タスクやカレンダーの同期後\n\n" +
    "📌 これにより、ストリークや学習メトリクスが最新の進捗を正確に反映します。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Metrics Troubleshooting Tooltip
function showDashboardErrorHelpTooltip() {
  const msg =
    "📉 TROUBLESHOOTING METRICS\n\n" +
    "- Make sure Dashboard cell B3 is set to a valid Phase (1–6).\n" +
    "- Check the 'Learning Log' sheet for missing or malformed data:\n" +
    "   • Date column (F) must contain valid dates.\n" +
    "   • Kanji/Vocab types should be logged correctly.\n" +
    "- Use 🩺 Heal Session Anchors if dates are off by one day.\n\n" +
    "📉 メトリクスのトラブルシューティング\n\n" +
    "- ダッシュボードのセル B3 に、有効なフェーズ（1〜6）が設定されていることを確認してください。\n" +
    "- 『Learning Log』シートで、以下のような不正データがないか確認してください：\n" +
    "   ・列 F（日付）が正しい日付形式であるか\n" +
    "   ・タイプが「Kanji」や「Vocabulary」など正しく入力されているか\n" +
    "- 日付が1日ずれている場合は、「🩺 Heal Session Anchors」を使用してください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Email Report Type Tooltip
function showEmailTypeTooltip() {
  const msg =
    "📬 EMAIL REPORT TYPES\n\n" +
    "🗓️ Daily Report:\n" +
    "- Quick snapshot of your latest streak, study time, and focus areas.\n\n" +
    "📈 Weekly Summary:\n" +
    "- A fuller breakdown of your study habits, resources used, and tag trends.\n\n" +
    "💡 Both reports pull directly from your Dashboard metrics.\n\n" +
    "📬 メールレポートの種類\n\n" +
    "🗓️ デイリーレポート：\n" +
    "- 直近のストリーク、学習時間、注力分野の簡易スナップショット。\n\n" +
    "📈 ウィークリーレポート：\n" +
    "- 学習習慣、使用リソース、タグの傾向などのより詳細な分析。\n\n" +
    "💡 どちらもダッシュボードのメトリクスを元に自動生成されます。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Email Troubleshooting Tooltip
function showEmailTroubleshootingTooltip() {
  const msg =
    "📧 EMAIL TROUBLESHOOTING\n\n" +
    "- Make sure your email is entered correctly in Dashboard cell B25.\n" +
    "- Check your spam folder just in case.\n" +
    "- Emails are only sent to your address and never stored or shared.\n\n" +
    "📧 メールのトラブルシューティング\n\n" +
    "- ダッシュボードのセル B25 に正しいメールアドレスが入力されているか確認してください。\n" +
    "- 念のため、迷惑メールフォルダもご確認ください。\n" +
    "- メールはご自身のアドレスのみに送信され、外部に保存・共有されることはありません。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Email Contents Tooltip
function showEmailContentsTooltip() {
  const msg =
    "📘 WHAT’S INCLUDED IN YOUR REPORT\n\n" +
    "- Study streak (🔥), longest streak, days studied\n" +
    "- Kanji/Vocab entries and total study time\n" +
    "- Top tags and most used study resource\n" +
    "- Average difficulty rating\n\n" +
    "📊 Pulled straight from your Dashboard for accurate summaries.\n\n" +
    "📘 レポートに含まれる情報\n\n" +
    "- ストリーク（🔥）、最長継続日数、学習日数\n" +
    "- 漢字・語彙エントリ数と合計学習時間\n" +
    "- 最も使用されたタグとリソース\n" +
    "- 難易度の平均スコア\n\n" +
    "📊 ダッシュボードの最新データをもとに正確に生成されます。";
  SpreadsheetApp.getUi().alert(msg);
}

// ✅ Study Logger Tooltip
function showStudyLoggerHelp() {
  const msg =
    "📝 STUDY LOGGER HELP\n\n" +
    "- Use this to log structured or meaningful study sessions.\n" +
    "- These entries are saved in the 'Learning Log' sheet.\n" +
    "- Always include *Time* (in minutes) and *Difficulty* (1–5) for full dashboard insights.\n\n" +
    "💡 FIELD TIPS:\n" +
    "• Type → Enter from study types (e.g., Shadowing, Lesson, Immersion).\n" +
    "• Tags → Comma-separated (e.g., beginner, N5, kana-heavy)\n" +
    "• Time → Total study time in minutes (e.g., 45)\n" +
    "• Difficulty → Rate from 1 (easy) to 5 (hard)\n\n" +
    "✅ Note: Entries with missing time or content won’t appear in metrics.\n\n" +
    "📝 学習記録ロガーの使い方\n\n" +
    "- 意味のある学習セッションを記録するためのツールです。\n" +
    "- 入力内容は「Learning Log（学習ログ）」シートに保存されます。\n" +
    "- *学習時間（分）* と *難易度（1〜5）* を必ず入力してください。\n\n" +
    "💡 各フィールドのヒント：\n" +
    "• Type → 学習タイプを入力（例：シャドーイング、レッスン、没入）\n" +
    "• Tags → カンマ区切り（例：beginner, N5, kana-heavy）\n" +
    "• Time → 合計学習時間（例：45分）\n" +
    "• Difficulty → 難易度を1（簡単）〜5（難しい）で評価\n\n" +
    "✅ メモ：学習時間や内容が入力されていない記録はメトリクスに反映されません。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Kanji and Vocab Logger Tooltip
function showKanjiVocabLoggerHelp() {
  const msg =
    "🔤 KANJI & VOCAB LOGGER HELP\n\n" +
    "- Quickly log Kanji or Vocabulary you've studied.\n" +
    "- These entries appear in your Dashboard under 'Kanji/Vocab Learned'.\n\n" +
    "💡 EXAMPLES:\n" +
    "• 水 (みず) = water\n" +
    "• 食べる (たべる) = to eat\n\n" +
    "💡 FIELD TIPS:\n" +
    "• Type → Choose 'Kanji' or 'Vocabulary'\n" +
    "• Tags → Optional, e.g., JLPT N5, core 2k\n\n" +
    "✅ Note: Both Character and Type are required for a valid entry.\n\n" +
    "🔤 漢字＆語彙ロガーの使い方\n\n" +
    "- 学習した漢字や語彙をすばやく記録するためのフォームです。\n" +
    "- 記録はダッシュボードの『Kanji/Vocab Learned』に反映されます。\n\n" +
    "💡 例：\n" +
    "• 水 (みず) = water\n" +
    "• 食べる (たべる) = to eat\n\n" +
    "💡 フィールドのヒント：\n" +
    "• Type → 「漢字」または「語彙」から選択\n" +
    "• Tags → 任意入力（例：JLPT N5、core 2k）\n\n" +
    "✅ 注意：文字（Character）とタイプ（Type）の両方が必要です。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Chunk Overview Tooltip
function showChunkOverviewTooltip() {
  const msg =
    "📦 WHAT ARE CHUNKS?\n\n" +
    "- Google Apps Script has API rate limits. To prevent sync failures, tasks and calendar events are split into 30-day blocks.\n\n" +
    "📊 CHUNK BREAKDOWN:\n" +
    "• Chunk 1 → Days 1–30\n" +
    "• Chunk 2 → Days 31–60\n" +
    "• Chunk 3 → Days 61–90\n" +
    "• Chunk 4 → Days 91–120\n" +
    "• Chunk 5 → Days 121–150\n" +
    "• Chunk 6 → Days 151–180\n\n" +
    "✅ You can sync chunks in any order, but only **after assigning dates**.\n\n" +
    "⚠️ Re-syncing the same chunk without wiping first can create duplicates.\n" +
    "Use the Clean-Up tools before repeating any chunk.\n\n" +
    "📦 チャンク（分割同期）とは？\n\n" +
    "- Google Apps Script には API 制限があります。そのため、同期エラーを防ぐためにタスクとカレンダーイベントは30日単位で分割されます。\n\n" +
    "📊 チャンクの構成：\n" +
    "• チャンク1 → 1〜30日\n" +
    "• チャンク2 → 31〜60日\n" +
    "• チャンク3 → 61〜90日\n" +
    "• チャンク4 → 91〜120日\n" +
    "• チャンク5 → 121〜150日\n" +
    "• チャンク6 → 151〜180日\n\n" +
    "✅ チャンクは任意の順番で同期可能ですが、**日付の割り当て後に行ってください**。\n\n" +
    "⚠️ 同じチャンクを再同期する前に、クリーンアップを行わないと重複が発生する可能性があります。\n" +
    "再同期の際は、必ずクリーンアップツールをご利用ください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Email Report Tooltip
function showEmailReportsTooltip() {
  const msg =
    "📧 EMAIL REPORTS\n\n" +
    "- Weekly Report: A summary of all key metrics is sent to the email in Dashboard cell B25.\n" +
    "- Daily Report: Sends a short snapshot of your most recent activity and streak.\n" +
    "- Before sending, the system updates your Dashboard to reflect latest stats.\n\n" +
    "⚠️ Make sure your email is correct and includes '@'.\n\n" +
    "✅ Tip: Reports can be used for journaling, accountability, or motivation!\n\n" +
    "📧 メールレポートについて\n\n" +
    "- 週間レポート：重要なメトリクスの概要がダッシュボードのセル B25 に記載されたメールアドレスに送信されます。\n" +
    "- 日次レポート：最新の学習活動と連続日数（ストリーク）のスナップショットを送信します。\n" +
    "- レポート送信前に、ダッシュボードが最新の状態に更新されます。\n\n" +
    "⚠️ メールアドレスが正しく入力されており、@ が含まれていることを確認してください。\n\n" +
    "✅ ヒント：レポートは学習記録、進捗管理、モチベーション維持に役立ちます。";
  SpreadsheetApp.getUi().alert(msg);
}

// ✅ Quick Study Logger Tooltip
function showQuickStudyLoggerTooltip() {
  const msg =
    "🈷️ QUICK STUDY LOGGER\n\n" +
    "- Use this tool to rapidly log brief entries: Kanji, Vocabulary, Grammar, or Immersion.\n" +
    "- Entries are saved directly to your active phase’s Learning Log.\n" +
    "- You MUST select a date to ensure proper timezone alignment.\n\n" +
    "✅ Ideal for:\n" +
    "• SRS practice (WaniKani, Anki, etc.)\n" +
    "• Passive immersion (Anime episodes, JP media)\n" +
    "• Small wins (e.g., 10-minute grammar drill)\n\n" +
    "🈷️ クイック学習ロガー\n\n" +
    "- このツールでは、漢字・語彙・文法・没入型学習などを素早く記録できます。\n" +
    "- 記録は、アクティブなフェーズの『Learning Log』に保存されます。\n" +
    "- タイムゾーンのずれを避けるために、日付は必ず手動で選択してください。\n\n" +
    "✅ 次のようなケースに最適：\n" +
    "・SRS練習（WaniKani、Ankiなど）\n" +
    "・受動的なインプット\n" +
    "・ちょっとした文法トレーニングなど";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Study Session Logger Tooltip
function showStudySessionLoggerTooltip() {
  const msg =
    "✏️ STUDY SESSION LOGGER\n\n" +
    "- Log detailed study sessions with time, difficulty, content, and tags.\n" +
    "- These entries populate the Learning Log and feed into your dashboard metrics.\n\n" +
    "✅ Pro Tip: Use consistent tags to see trends in your dashboard (e.g., 'immersion', 'shadowing').\n\n" +
    "✏️ 学習セッション記録フォーム\n\n" +
    "- 学習時間、難易度、内容、タグを入力して、詳細なセッションを記録できます。\n" +
    "- 記録は「Learning Log」に反映され、ダッシュボードのメトリクスに反映されます。\n\n" +
    "✅ プロのヒント：タグを統一して使うと、ダッシュボードで学習傾向が見えやすくなります（例：「immersion」「shadowing」など）。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Dashboard Styling Tooltip
function showDashboardStylingTooltip() {
  const msg =
    "🎨 DASHBOARD STYLING\n\n" +
    "- Automatically formats your Dashboard based on the selected theme.\n" +
    "- Applies clean layout, section colors, font styling, and streak highlights.\n\n" +
    "✅ Tip: Use 'Apply Theme to All Sheets' and/or 'Format Dashboard' if the dashboard looks broken or off-alignment.\n\n" +
    "🎨 ダッシュボードのスタイリング\n\n" +
    "- 選択されたテーマに基づいて、ダッシュボードのレイアウトが自動的に整えられます。\n" +
    "- セクションの色、フォントスタイル、ストリークのハイライトなどが適用されます。\n\n" +
    "✅ ヒント：「Apply Theme to All Sheets」や「Format Dashboard（ダッシュボード整形）」を使うと、レイアウトの崩れやズレを修正できます。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Sheet Styling Tooltip
function showSheetStylingTooltip() {
  const msg =
    "🧽 SHEET STYLING SYSTEM\n\n" +
    "- Applies consistent fonts, spacing, striping, and sizing across all sheets.\n" +
    "- Ensures each page is clean, readable, and easy to scan.\n\n" +
    "✅ Tip: Use this if your Learning Log or Master Sheet looks messy or misaligned.\n\n" +
    "🧽 シートのスタイル整形システム\n\n" +
    "- すべてのシートに一貫したフォント、間隔、ストライプ、サイズ設定を適用します。\n" +
    "- 各ページが見やすく、読みやすく整えられます。\n\n" +
    "✅ ヒント：「Learning Log」や「Master Sheet」の表示が乱れているときに使用してください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Themes Tooltip
function showThemeTooltip() {
  const msg =
    "🎨 THEME SYSTEM\n\n" +
    "- Choose between Light and Dark themes using the sidebar.\n" +
    "- This updates the Dashboard styling and text colors.\n" +
    "- Click 'Apply to All Sheets' to apply the style workbook-wide.\n\n" +
    "💡 Your selection is saved in Dashboard cell B30.\n\n" +
    "🎨 テーマシステム\n\n" +
    "- サイドバーでライトテーマまたはダークテーマを選択できます。\n" +
    "- 選択内容に応じて、ダッシュボードのスタイルと文字色が更新されます。\n" +
    "- 「Apply to All Sheets」をクリックすると、ブック全体にテーマが適用されます。\n\n" +
    "💡 選択されたテーマは、ダッシュボードのセル B30 に保存されます。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Heal Session Date Anchors Tooltip
function showHealTooltip() {
  const msg =
    "🩺 HEAL SESSION ANCHORS\n\n" +
    "- Google saves dates in UTC, which causes some logs to show the **previous day** if you're behind UTC.\n" +
    "- This tool fixes that by converting the dates into clean text using your **local timezone**.\n\n" +
    "✅ Run this after logging sessions to ensure accurate display.\n" +
    "⚠️ Required if you live in the US or anywhere behind UTC.\n\n" +
    "🩺 セッションアンカーの修復\n\n" +
    "- Google は日付を **UTC（協定世界時）** で保存するため、\n" +
    "　UTCより遅れているタイムゾーン（例：米国）では、前日として表示されることがあります。\n" +
    "- このツールは、**ローカルタイムゾーン** に基づいた日付に変換し、表示のズレを修正します。\n\n" +
    "✅ セッションを記録した後にこのツールを実行することで、日付が正しく表示されます。\n" +
    "⚠️ 米国など、UTCより遅れている地域に住んでいる方には必須のステップです。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Create Task Lists Tooltip
function showCreateTaskListTooltip() {
  const msg =
    "🧾 CREATE TASK LISTS\n\n" +
    "- This sets up the Google Task Lists required for syncing your daily lessons.\n" +
    "- Each phase (1–6) gets its own task list, linked to your Master Sheet.\n" +
    "- You MUST run this **before syncing any tasks** — otherwise syncs will fail.\n\n" +
    "✅ Use after assigning dates, then run task sync.\n\n" +
    "🧾 タスクリストの作成\n\n" +
    "- 日次レッスンを同期するために必要な Google タスクリストを作成します。\n" +
    "- 各フェーズ（1〜6）に対応するリストが作られ、Master Sheet と連携します。\n" +
    "- **タスクを同期する前に必ず実行してください**。\n\n" +
    "✅ 日付の割り当て後に実行し、その後タスクの同期を行ってください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Test Sheet and Calendar Connection Tooltip
function showTestConnectionTooltip() {
  const msg =
    "🧪 TEST CONNECTION TO SHEET OR CALENDAR\n\n" +
    "- These tools help confirm that your Sheet ID (B27) and Calendar ID (B28) are valid and reachable.\n" +
    "- Use them if sync is failing unexpectedly.\n\n" +
    "🧪 シートまたはカレンダーとの接続テスト\n\n" +
    "- このツールを使って、B27（シートID）とB28（カレンダーID）が正しく設定されているか確認できます。\n" +
    "- 同期に問題がある場合は、まずこのテストを行ってください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Smart Prompt Tooltip
function showSmartPromptTooltip() {
  const msg =
    "🎲 SMARTPROMPT STUDY SUGGESTIONS\n\n" +
    "- This tool gives you randomized study suggestions based on your current mood.\n" +
    "- Select a mood (Low, Normal, High) and click the button to receive an idea for how to study today.\n" +
    "- It’s designed for inspiration — especially useful when you feel stuck or indecisive.\n\n" +
    "✅ Pro Tip: If the first one doesn’t click, click again!\n\n" +
    "🎲 スマートプロンプト（学習提案ツール）\n\n" +
    "- 今の気分に基づいて、ランダムな学習アイデアを提案します。\n" +
    "- 気分（Low、Normal、High）を選んでボタンを押すだけ。\n" +
    "- 学習に迷ったとき、やる気が出ないときに最適なインスピレーション源です。\n\n" +
    "✅ ヒント：ピンとこない提案が出たら、もう一度クリックしてみましょう！";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Quick vs Full Session Entry Tooltip
function showQuickEntryHelpTooltip() {
  const msg =
    "📝 QUICK ENTRY VS FULL SESSION\n\n" +
    "- **Quick Entry** → For fast tracking: just Type, Content, Tags, and Date.\n" +
    "- **Full Session** → Captures Time, Difficulty, Completion status, Notes, and more.\n\n" +
    "📘 Use Quick Entry when:\n" +
    "• You’re logging something short or informal\n" +
    "• You want speed and simplicity\n\n" +
    "📘 Use Full Session when:\n" +
    "• You’re doing focused study with tracked effort\n" +
    "• You want metrics like duration or difficulty included\n\n" +
    "✅ Both feed into your Learning Log and dashboard insights.\n\n" +
    "📝 クイックエントリー vs フルセッション\n\n" +
    "- **クイックエントリー** → Type・内容・タグ・日付のみを記録（スピーディ）\n" +
    "- **フルセッション** → 学習時間・難易度・完了状況・メモなど詳細を記録\n\n" +
    "📘 クイックエントリーのおすすめ場面：\n" +
    "・短時間・軽めの学習記録\n" +
    "・素早く入力したいとき\n\n" +
    "📘 フルセッションのおすすめ場面：\n" +
    "・集中して行った学習を記録したいとき\n" +
    "・時間や難易度などのデータも残したいとき\n\n" +
    "✅ どちらの形式でも『Learning Log』とダッシュボードに反映されます。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ New Phase Log Tooltip
function showNewPhaseLogTooltip() {
  const msg =
    "📄 NEW PHASE LOG SHEETS\n\n" +
    "- When you change the 'Current Phase' (Dashboard cell B3) and click '🧠 Update Dashboard', the system auto-generates a new Learning Log for that phase.\n" +
    "- This prevents bloat in the original log and keeps sessions cleanly organized.\n\n" +
    "✅ Sheet names follow the pattern: 'Learning Log', 'Learning Log 2', 'Learning Log 3', etc.\n" +
    "⚠️ Each new phase must have a start date (B4) entered before updating.\n\n" +
    "📄 新しいフェーズ用ログシート\n\n" +
    "- 「Dashboard」のセル B3（Current Phase）を変更し、「🧠 Update Dashboard」を実行すると、\n" +
    "　そのフェーズ用の新しい『Learning Log』シートが自動生成されます。\n" +
    "- これにより、元のログが肥大化するのを防ぎ、セッションをフェーズごとに整理できます。\n\n" +
    "✅ シート名の例：'Learning Log'、'Learning Log 2'、'Learning Log 3' …\n" +
    "⚠️ フェーズ開始日（セル B4）を設定してから実行してください。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Extend Phase Dates Tooltip
function showExtendPhaseDatesTooltip() {
  const msg =
    "📈 EXTEND PHASE DATES\n\n" +
    "- Adds 30 additional days to the current Phase block in your Master Lesson Plan.\n" +
    "- Uses the last assigned date of the Phase as the new starting point.\n" +
    "- Maintains continuity — perfect if your study pace is slower than expected or you want to expand the curriculum.\n\n" +
    "⚠️ Make sure your Phase (Dashboard B3) and Phase Start Date (B4) are correctly set.\n" +
    "📌 Only affects the currently selected Phase — other phases remain unchanged.\n\n" +
    "📈 フェーズ日付の延長\n\n" +
    "- 現在のフェーズに、Master Lesson Planでさらに30日分を追加します。\n" +
    "- フェーズの最後に割り当てられた日付を起点として新しい日付を追加します。\n" +
    "- 学習ペースが遅れている場合や、カリキュラムを拡張したい場合に最適です。\n\n" +
    "⚠️ フェーズ（ダッシュボードB3）と開始日（B4）が正しく設定されていることを確認してください。\n" +
    "📌 他のフェーズには影響しません。選択中のフェーズのみ変更されます。";
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Type/Tag Reference Tooltip
function showTooltip_TypeTagReference() {
  const msg =
    "🔖 TYPE TAG\n\n" +
    "- Each entry in the Learning Log is automatically categorized by study type.\n" +
    "- These tags help organize your data visually and support dashboard filtering.\n" +
    "- Tags are derived from the 'Type' column (Column B).\n\n" +
    "Current Type Tags:\n" +
    "• 📚 Vocabulary — Vocabulary-focused entries (word lists, flashcards, sentence building)\n" +
    "• 🧠 Grammar — Grammar explanations, conjugations, and pattern studies\n" +
    "• 🎧 Immersion — Passive or active media (anime, podcasts, books, etc.)\n" +
    "• 📝 Study Session — Structured study time (textbooks, workbooks)\n" +
    "• 🎬 Season Entry — Full anime/TV seasons; logs total time but not session average\n\n" +
    "🗂️ Filter by Type Tag to isolate different learning styles or content types.\n\n" +
    "🔖 分類タイプ\n\n" +
    "- Learning Log の各エントリは、自動的に学習タイプ別に分類されます。\n" +
    "- タグは視覚的整理やダッシュボードでのフィルタリングに役立ちます。\n" +
    "- 「Type」列（列B）の値から自動的に取得されます。\n\n" +
    "現在のタイプ分類:\n" +
    "• 📚 Vocabulary（語彙）— 単語リスト、フラッシュカード、例文構築など\n" +
    "• 🧠 Grammar（文法）— 文法解説、活用、構文パターンの学習\n" +
    "• 🎧 Immersion（没入）— アニメ、ポッドキャスト、本などの受動的/能動的なインプット\n" +
    "• 📝 Study Session（学習セッション）— 教科書やワークブックによる集中学習\n" +
    "• 🎬 Season Entry（シリーズ視聴）— アニメやドラマのシーズン全体の学習時間を記録（平均時間には含まれません）\n\n" +
    "🗂️ Type Tag でフィルターすると、異なる学習スタイルや教材タイプを簡単に整理できます。";
    
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Season Logger Tooltip
function getTooltip_SeasonLogger() {
  const msg =
    "🎬 SEASON LOGGER\n\n" +
    "- Use this to log a full anime or TV show season as a single immersive session.\n" +
    "- These entries contribute to total study time, but are excluded from session average calculations.\n" +
    "- Ideal for long-form media like full seasons or documentary marathons.\n\n" +
    "🎞️ What about movies?\n" +
    "- We recommend logging movies (usually 1.5–2.5 hours) using the Quick Logger or Study Session logger.\n" +
    "- This allows them to count toward your average session time as true immersive learning.\n" +
    "- BUT — if you prefer to exclude *all* media from your session averages, feel free to log movies using the Season Logger too!\n\n" +
    "⏱️ Logged Type: Season Entry\n" +
    "✅ Adds to total study hours\n" +
    "🚫 Excluded from average session time\n\n" +
    "🎬 シーズンロガー\n\n" +
    "- アニメやドラマのシーズン全体を1つの没入型セッションとして記録できます。\n" +
    "- 記録された時間は「総学習時間」に加算されますが、「平均セッション時間」には含まれません。\n" +
    "- シーズン作品や長編ドキュメンタリーなどに最適です。\n\n" +
    "🎞️ 映画はどうする？\n" +
    "- 映画（通常1.5～2.5時間）は「Quick Logger」または「Study Session Logger」で記録するのがおすすめです。\n" +
    "- これにより、平均セッション時間にカウントされ、没入型学習として反映されます。\n" +
    "- ただし、すべてのメディアを平均時間から除外したい場合は、映画も「Season Logger」で記録してOKです！\n\n" +
    "⏱️ ログタイプ：Season Entry\n" +
    "✅ 総学習時間に加算されます\n" +
    "🚫 平均セッション時間には含まれません";
  
  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Refresh Type Tag Tooltip
function showTooltip_RefreshTypeTagColumn() {
  const msg =
    "🔁 REFRESH TYPE TAGS (COLUMN K)\n\n" +
    "- Reapplies the Type Tag formula to match Column B (Type)\n" +
    "- Ensures all entries are properly tagged for filtering and dashboard grouping\n" +
    "- Especially useful after manual edits, copy-pasting, or if some rows show blank tags\n\n" +
    "📌 This is a visual helper column only — it does not affect study totals or metrics.\n\n" +
    "📎 When to use:\n" +
    "- You've added new rows manually\n" +
    "- Some tags appear empty\n" +
    "- You copied/pasted and broke the formula chain\n\n" +
    "———\n\n" +
    "🔁 タイプタグの更新（列K）\n\n" +
    "- 列B（Type）の内容をもとに、列Kの「Type Tag」列を再計算して反映します。\n" +
    "- すべてのエントリが正しく分類され、フィルターやダッシュボードでの表示に役立ちます。\n" +
    "- 手動追加やコピー＆ペースト後、タグが欠落している場合などに特に有効です。\n\n" +
    "📌 この列は視覚的な補助であり、学習時間やメトリクスには影響しません。\n\n" +
    "📎 使用タイミング：\n" +
    "- 行を手動で追加した場合\n" +
    "- 一部のタグが空欄の場合\n" +
    "- 数式チェーンが壊れた場合（コピー＆ペースト後など）";

  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Real OAuth Tooltip
function showTooltip_RealOAuth() {
  const msg =
    "🔐 REAL OAUTH\n\n" +
    "- Re-triggers Google’s internal OAuth permissions handshake\n" +
    "- Fixes issues where scripts silently fail or sidebar tools won’t open\n" +
    "- Can resolve 'Authorization Required' errors without needing full redeploy\n\n" +
    "🛠 How to use:\n" +
    "1. Run the 🧭 Real OAuth tool from the Debug Tools menu\n" +
    "2. Log out of your Google Account completely\n" +
    "3. Log back in and reload the sheet\n\n" +
    "📌 Notes:\n" +
    "- Google may or may not prompt the OAuth popup — it's inconsistent\n" +
    "- This method forces a read + write event to reset stuck auth tokens\n" +
    "- It's a known quirk of Google Sheets scripting — not user error\n\n" +
    "📎 Use this if:\n" +
    "- Sidebar tools won’t load\n" +
    "- Loggers silently fail to submit\n" +
    "- You're getting unexpected permission errors\n\n" +
    "———\n\n" +
    "🔐 リアルOAuth（再認証の強制）\n\n" +
    "- Googleの内部OAuth認証を再トリガーして再許可を促します\n" +
    "- スクリプトが動かなくなったり、サイドバーが開かない不具合を解消できます\n" +
    "- 「認証が必要です」といったエラーを、再デプロイせずに解決する方法です\n\n" +
    "🛠 使用方法：\n" +
    "1. メニュー「🧭 Debug Tools」から「Real OAuth」を実行します\n" +
    "2. Googleアカウントから完全にログアウトします\n" +
    "3. 再ログインしてシートを再読み込みします\n\n" +
    "📌 補足：\n" +
    "- GoogleがOAuth承認画面を表示するかは日によって異なります（気まぐれです）\n" +
    "- 読み取りと書き込み操作を強制実行することで、認証のリフレッシュを行います\n" +
    "- これはGoogle Sheetsスクリプトの仕様によるもので、操作ミスではありません\n\n" +
    "📎 このツールを使うべき場面：\n" +
    "- サイドバーが開かない\n" +
    "- ログ送信が無反応になる\n" +
    "- 許可が必要とのエラーが出る";

  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Format Dashboard Tooltip
function showTooltip_FormatDashboard() {
  const msg =
    "🧼 FORMAT DASHBOARD\n\n" +
    "- Reapplies the default formatting and color scheme to the Dashboard sheet\n" +
    "- Fixes layout issues if you've accidentally altered row heights, colors, merged cells, or font alignment\n" +
    "- Pulls the Light or Dark theme based on cell B30 and reapplies all header blocks, zones, and labels\n\n" +
    "📌 Does not change any formulas, data, or metrics — purely visual.\n\n" +
    "📎 When to use:\n" +
    "- Dashboard layout looks broken or inconsistent\n" +
    "- You deleted cell notes or formatting\n" +
    "- You switched themes and things didn’t update cleanly\n\n" +
    "———\n\n" +
    "🧼 ダッシュボードのフォーマット修正\n\n" +
    "- ダッシュボードの初期レイアウトと配色テーマを再適用します（ライト/ダークはB30に基づきます）\n" +
    "- 行の高さ、色、結合セル、フォント配置などをリセットして整えます\n" +
    "- ヘッダー、フェーズゾーン、メトリクス表示、メール/IDブロックなどを完全に再フォーマットします\n\n" +
    "📌 数式やデータは変更されません。見た目だけの修正です。\n\n" +
    "📎 使用タイミング：\n" +
    "- ダッシュボードの見た目が崩れたとき\n" +
    "- セルのメモや装飾を消してしまったとき\n" +
    "- テーマを切り替えたのに表示が正しく反映されなかったとき";

  SpreadsheetApp.getUi().alert(msg);
}


// ✅ Trim Learning Log Tooltip
function showTrimLearningLogTooltip() {
  const msg =
    "✂️ MANUAL TRIM: LEARNING LOG\n\n" +
    "- This tool removes unused blank rows at the bottom of the Learning Log.\n" +
    "- It should NOT be necessary under normal circumstances — our logger scripts\n" +
    "  automatically manage row trimming to prevent logging issues.\n" +
    "- However, when changing themes or due to occasional Google Sheets quirks,\n" +
    "  extra rows may sometimes be re-added unintentionally.\n\n" +
    "✅ Use this function only if you see your Learning Log with 1000+ blank rows.\n\n" +
    "✂️ 学習ログの手動トリム\n\n" +
    "- このツールは、学習ログシートの下部にある未使用の空白行を削除します。\n" +
    "- 通常は必要ありません。ロガースクリプトが自動的に行を管理しています。\n" +
    "- ただし、テーマの変更や Google Sheets の仕様によって、\n" +
    "  不要な行が再度追加されることがあります。\n\n" +
    "✅ 学習ログに 1000 行以上の空白がある場合のみ実行してください。";
  SpreadsheetApp.getUi().alert(msg);
}

