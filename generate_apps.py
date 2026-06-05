import pandas as pd
import json
import os
import re

excel_file = "見出語・用例リスト(Part 1～5，＋α，外来語).xlsx"
xls = pd.ExcelFile(excel_file)

with open("leap_app.html", "r", encoding="utf-8") as f:
    template_html = f.read()

part1_html = template_html.split("const chapterData = [")[0] + "const chapterData = "
part2_html = template_html.split("];\n\n        let activeSentenceElement")[1]
part2_html = ";\n\n        let activeSentenceElement" + part2_html

# Fix Title
part1_html = part1_html.replace("LEAP 学習アプリ", "LEAP Basic学習アプリ")

new_format_func = """
        function formatEnglishSentence(item) {
            const enSentence = item.en;
            const answer = item.answer;
            const blank = '__________';

            if (!enSentence) return `<span class="answer-span text-3xl">${answer}</span>`;

            let baseWord = answer.replace(/[～AB]/g, '').replace(/（.*?）/g, '').replace(/\\(.*?\\)/g, '').trim();
            let searchWords = [];
            
            if (baseWord.length > 0) searchWords.push(baseWord);
            
            if (baseWord === 'oneself') {
                searchWords.push('myself', 'yourself', 'himself', 'herself', 'itself', 'ourselves', 'yourselves', 'themselves');
            } else {
                const words = baseWord.split(/\\s+/).filter(w => w.length > 0 && /^[a-zA-Z]+$/.test(w));
                if (words.length > 0) {
                    let w = words[0].toLowerCase();
                    searchWords.push(w);
                    // Add regular forms
                    searchWords.push(w + 's', w + 'es', w + 'd', w + 'ed', w + 'ing');
                    if (w.endsWith('e')) searchWords.push(w.slice(0, -1) + 'ing');
                    if (w.endsWith('y')) searchWords.push(w.slice(0, -1) + 'ies', w.slice(0, -1) + 'ied');
                    if (w.length > 2) {
                        const last = w.slice(-1);
                        searchWords.push(w + last + 'ed', w + last + 'ing');
                    }
                    
                    const irregulars = {
                        'write': ['wrote', 'written'], 'speak': ['spoke', 'spoken'], 'say': ['said'], 'tell': ['told'],
                        'see': ['saw', 'seen'], 'hear': ['heard'], 'take': ['took', 'taken'], 'meet': ['met'],
                        'feel': ['felt'], 'think': ['thought'], 'know': ['knew', 'known'], 'choose': ['chose', 'chosen'],
                        'leave': ['left'], 'make': ['made'], 'catch': ['caught'], 'find': ['found'], 'buy': ['bought'],
                        'bring': ['brought'], 'teach': ['taught'], 'understand': ['understood'], 'build': ['built'],
                        'send': ['sent'], 'spend': ['spent'], 'lose': ['lost'], 'sleep': ['slept'], 'keep': ['kept'],
                        'do': ['did', 'done'], 'go': ['went', 'gone'], 'come': ['came'], 'run': ['ran'],
                        'give': ['gave', 'given'], 'get': ['got', 'gotten'], 'forget': ['forgot', 'forgotten'],
                        'begin': ['began', 'begun'], 'drink': ['drank', 'drunk'], 'swim': ['swam', 'swum'],
                        'sing': ['sang', 'sung'], 'ring': ['rang', 'rung'], 'drive': ['drove', 'driven'],
                        'ride': ['rode', 'ridden'], 'eat': ['ate', 'eaten'], 'fall': ['fell', 'fallen'],
                        'hold': ['held'], 'stand': ['stood'], 'sit': ['sat'], 'win': ['won'], 'pay': ['paid'],
                        'sell': ['sold'], 'break': ['broke', 'broken'], 'wear': ['wore', 'worn'],
                        'grow': ['grew', 'grown'], 'fly': ['flew', 'flown'], 'blow': ['blew', 'blown'],
                        'draw': ['drew', 'drawn'], 'throw': ['threw', 'thrown'], 'hide': ['hid', 'hidden'],
                        'fight': ['fought'], 'feed': ['fed'], 'strike': ['struck'], 'wake': ['woke', 'woken'],
                        'bear': ['bore', 'born'], 'beat': ['beat', 'beaten'], 'bind': ['bound'], 'bleed': ['bled'],
                        'breed': ['bred'], 'tear': ['tore', 'torn'], 'sweep': ['swept'], 'weep': ['wept'],
                        'bend': ['bent'], 'lend': ['lent'], 'mean': ['meant'], 'lead': ['led'], 'shoot': ['shot']
                    };
                    
                    if (irregulars[w]) searchWords.push(...irregulars[w]);
                }
            }

            // We only want to replace the FIRST matching word so we don't mess up multiple instances of the same word unless we use replaceAll.
            // Actually replaceAll with regex /g is better if the word appears multiple times.
            for (let actualWord of searchWords) {
                if (!actualWord) continue;
                try {
                    const regex = new RegExp(`\\\\b${actualWord}\\\\b`, 'gi');
                    if (regex.test(enSentence)) {
                        return enSentence.replace(regex, match => `<span class="answer-span">${match}</span><span class="underline-span">${blank}</span>`);
                    }
                } catch (e) { }
            }
            
            // Substring fallback
            for (let actualWord of searchWords) {
                if (!actualWord) continue;
                if (enSentence.toLowerCase().includes(actualWord.toLowerCase())) {
                    const index = enSentence.toLowerCase().indexOf(actualWord.toLowerCase());
                    if (index !== -1) {
                        const originalWord = enSentence.substring(index, index + actualWord.length);
                        return enSentence.substring(0, index) +
                            `<span class="answer-span">${originalWord}</span><span class="underline-span">${blank}</span>` +
                            enSentence.substring(index + actualWord.length);
                    }
                }
            }

            return enSentence;
        }
"""

part2_html = re.sub(
    r'function formatEnglishSentence\(item\) \{.*?\n        \}',
    lambda m: new_format_func.strip(),
    part2_html,
    flags=re.DOTALL
)

new_start_quiz = """
        function startQuiz() {
            const questions = [...chapterData].sort(() => 0.5 - Math.random()).slice(0, 10);
            let currentIdx = 0, score = 0;

            function extractBlankWord(item) {
                const formatted = formatEnglishSentence(item);
                const match = formatted.match(/<span class="answer-span">([\\s\\S]*?)<\\/span>/);
                return match ? match[1].trim() : item.answer;
            }

            function nextQuestion() {
                if (currentIdx >= questions.length) {
                    const finalScore = Math.round(score / questions.length * 100);
                    if (window.saveLeapScore) {
                        window.saveLeapScore('quizScore', finalScore);
                    } else {
                        console.error("Firebase saveLeapScore is not loaded yet.");
                        alert("データベース連携が初期化されていません。ネットワーク接続が正常か、または広告ブロック/セキュリティ拡張機能により連携スクリプトの読込がブロックされていないかご確認ください。");
                    }
                    quizContent.innerHTML = `<h3 class="text-3xl font-bold mb-4">結果: ${score}/${questions.length}</h3><p class="text-6xl text-purple-600 font-bold">${finalScore}点</p>`;
                    return;
                }
                const q = questions[currentIdx];
                const correctWord = extractBlankWord(q);
                
                const formattedQ = formatEnglishSentence(q);
                let displaySentence = formattedQ.replace(/<span class="answer-span">[\\s\\S]*?<\\/span><span class="underline-span">__________<\\/span>/g, '__________');
                if (displaySentence === q.en && q.en) {
                    displaySentence = q.en;
                } else if (!q.en) {
                    displaySentence = "(例文なし)";
                }

                const allOtherWords = chapterData.filter(d => d.id !== q.id).map(d => extractBlankWord(d));
                const uniqueOtherWords = [...new Set(allOtherWords)].filter(w => w !== correctWord);
                const distractors = uniqueOtherWords.sort(() => 0.5 - Math.random()).slice(0, 3);
                
                const options = [correctWord, ...distractors].sort(() => 0.5 - Math.random());

                quizContent.innerHTML = `
            <p class="mb-2 text-gray-500">問題 ${currentIdx + 1}/10</p>
            <p class="text-lg font-bold mb-2 text-gray-700">${q.ja}</p>
            <p class="text-xl font-bold mb-6 font-english text-blue-800">${displaySentence}</p>
            <div class="grid grid-cols-1 gap-3">
                ${options.map(opt => `<button class="quiz-opt p-3 border-2 border-gray-200 rounded-lg hover:bg-blue-50 hover:border-blue-300 font-english transition-colors">${opt}</button>`).join('')}
            </div>
        `;
                document.querySelectorAll('.quiz-opt').forEach(btn => {
                    btn.onclick = () => {
                        if (btn.innerText === correctWord) { btn.classList.add('bg-green-200', 'border-green-400'); score++; }
                        else { btn.classList.add('bg-red-200', 'border-red-400'); }
                        
                        // Disable buttons
                        document.querySelectorAll('.quiz-opt').forEach(b => b.disabled = true);
                        
                        setTimeout(() => { currentIdx++; nextQuestion(); }, 1000);
                    };
                });
            }
            quizModal.classList.remove('hidden');
            nextQuestion();
        }
"""

part2_html = re.sub(
    r'function startQuiz\(\) \{.*?\n        \}',
    lambda m: new_start_quiz.strip(),
    part2_html,
    flags=re.DOTALL
)

# Fix item.ja display
part2_html = part2_html.replace(
    '<p class="mb-2 text-gray-700">${item.ja}</p>',
    '${item.ja ? `<p class="mb-2 text-gray-700">${item.ja}</p>` : ""}'
)


# Intercept pronunciation score
part2_html = part2_html.replace(
    "showMessage(`",
    "if(window.saveLeapScore) window.saveLeapScore('pronunciationScore_' + activeSentenceElement.dataset.id, score);\n            showMessage(`"
)

new_pdf_logic1 = """
        document.getElementById('download-sentence-quiz-btn').onclick = () => {
            const items = getSelected();
            const html = `<ol>${items.map(i => {
                if (!i.en) return `<li>（例文なし: ${i.answer}）</li>`;
                let formatted = formatEnglishSentence(i);
                let displaySentence = formatted.replace(/<span class="answer-span">[\\s\\S]*?<\\/span><span class="underline-span">__________<\\/span>/g, '__________');
                if (displaySentence === i.en) displaySentence = i.en;
                return `<li>${displaySentence}<br><small>${i.ja}</small></li>`;
            }).join('')}</ol><div class="ans"><h3>解答</h3><p>${items.map(i => i.answer).join(', ')}</p></div>`;
            openPrintWindow(html, '例文穴埋めテスト');
        };
"""

new_pdf_logic2 = """
        document.getElementById('download-sentence-mixed-quiz-btn').onclick = () => {
            const items = getSelected();
            const html = `<ol>${items.map(i => {
                if (!i.en) {
                    return `<li>[英語]: ${i.answer}<br>[意味]: __________</li>`;
                }
                if (Math.random() > 0.5) {
                    let formatted = formatEnglishSentence(i);
                    let displaySentence = formatted.replace(/<span class="answer-span">[\\s\\S]*?<\\/span><span class="underline-span">__________<\\/span>/g, '__________');
                    if (displaySentence === i.en) displaySentence = i.en;
                    return `<li>${displaySentence}<br><small>${i.ja}</small></li>`;
                } else {
                    return `<li>[英語]: ${i.answer}<br>[意味]: __________</li>`;
                }
            }).join('')}</ol><div class="ans"><h3>解答</h3><p>${items.map(i => i.answer + ' / ' + i.explanation.split('\\n')[0]).join('<br>')}</p></div>`;
            openPrintWindow(html, 'ランダム混合テスト');
        };
"""

part2_html = re.sub(
    r"document\.getElementById\('download-sentence-quiz-btn'\)\.onclick = \(\) => \{.*?\n        \};",
    lambda m: new_pdf_logic1.strip(),
    part2_html,
    flags=re.DOTALL
)

part2_html = re.sub(
    r"document\.getElementById\('download-sentence-mixed-quiz-btn'\)\.onclick = \(\) => \{.*?\n        \};",
    lambda m: new_pdf_logic2.strip(),
    part2_html,
    flags=re.DOTALL
)

firebase_script = """
    <script type="module">
        import { initializeApp } from "https://www.gstatic.com/firebasejs/10.9.0/firebase-app.js";
        import { getAuth, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/10.9.0/firebase-auth.js";
        import { getFirestore, doc, setDoc, getDoc } from "https://www.gstatic.com/firebasejs/10.9.0/firebase-firestore.js";

        const firebaseConfig = {
            apiKey: "AIzaSyCNrvDMnlN_geijgRguQrolnTIZ_rdZwyw",
            authDomain: "sokudoku-stories-db.firebaseapp.com",
            projectId: "sokudoku-stories-db",
            storageBucket: "sokudoku-stories-db.firebasestorage.app",
            messagingSenderId: "979353290932",
            appId: "1:979353290932:web:f0c512cfbe64a7afe36fe1"
        };

        const app = initializeApp(firebaseConfig);
        const auth = getAuth(app);
        const db = getFirestore(app);

        const appId = window.location.pathname.split('/').pop().replace('.html', '');

        window.localMasteredWords = {};

        // ロード時にマスターデータを反映する
        window.applyLoadedMasteredWords = (masteredWords) => {
            window.localMasteredWords = masteredWords || {};
            document.querySelectorAll('.sentence-item').forEach(wrapper => {
                const wordId = wrapper.dataset.id;
                const btn = wrapper.querySelector('.toggle-master-btn-local');
                if (window.localMasteredWords[wordId] === true) {
                    wrapper.classList.add('word-mastered');
                    if(btn) btn.innerText = "⭐ マスター解除";
                } else {
                    wrapper.classList.remove('word-mastered');
                    if(btn) btn.innerText = "⭐ マスター";
                }
            });
        };

        // Firebaseにトグル状態を保存する
        window.toggleWordMasterState = (wordId, makeMastered) => {
            if (window.currentUser) {
                const docRef = doc(db, "leap_users", window.currentUser.uid, "progress", appId);
                window.localMasteredWords[wordId] = makeMastered;
                
                // ネストされたオブジェクトを直接渡す。merge: true により、Firestoreが既存のキーとマージしてくれます。
                const updateData = {
                    masteredWords: {
                        [wordId]: makeMastered
                    }
                };
                
                setDoc(docRef, updateData, { merge: true })
                .then(() => {
                    console.log(`Word ${wordId} master state:`, makeMastered);
                    window.applyLoadedMasteredWords(window.localMasteredWords);
                })
                .catch(e => {
                    console.error("Error updating word master:", e);
                    alert("マスター状態の保存に失敗しました: " + e.message);
                });
            } else {
                window.localMasteredWords[wordId] = makeMastered;
                window.applyLoadedMasteredWords(window.localMasteredWords);
                console.warn("Saving locally only.");
            }
        };

        onAuthStateChanged(auth, (user) => {
            if (user) {
                window.currentUser = user;
                console.log("Firebase Authenticated as:", user.email);
                const docRef = doc(db, "leap_users", user.uid, "progress", appId);
                
                // Firestoreからマスター済み単語データをロード
                getDoc(docRef).then((docSnap) => {
                    if (docSnap.exists()) {
                        const data = docSnap.data();
                        window.applyLoadedMasteredWords(data.masteredWords);
                    }
                }).catch(e => console.error("Error loading mastered words:", e));

                setDoc(docRef, {
                    started: true,
                    lastAccessed: new Date().toISOString()
                }, { merge: true }).then(() => {
                    console.log("Progress started state saved successfully.");
                }).catch(e => {
                    console.error("Progress save error:", e);
                    alert("進捗の自動保存に失敗しました: " + e.message);
                });
                
                setDoc(doc(db, "leap_users", user.uid), {
                    displayName: user.displayName,
                    email: user.email,
                    photoURL: user.photoURL,
                    lastLogin: new Date().toISOString()
                }, { merge: true }).catch(e => console.error("User profile save error:", e));
            } else {
                console.warn("No user session active on this page.");
            }
        });

        window.saveLeapScore = (type, score) => {
            if (window.currentUser) {
                const docRef = doc(db, "leap_users", window.currentUser.uid, "progress", appId);
                const updateData = {};
                updateData[type] = score;
                setDoc(docRef, updateData, { merge: true })
                .then(() => {
                    console.log("Score saved successfully:", type, score);
                })
                .catch(e => {
                    console.error("Score save error:", e);
                    alert("テスト結果の保存に失敗しました: " + e.message);
                });
            } else {
                console.error("Score save failed: No user is logged in.");
                alert("ログイン状態が確認できないため、テスト結果を保存できませんでした。ダッシュボードから再度ログインしてください。");
            }
        };
    </script>
"""

def rreplace(s, old, new, occurrence):
    li = s.rsplit(old, occurrence)
    return new.join(li)

part2_html = rreplace(part2_html, "</body>", firebase_script + "\n</body>", 1)

def clean_english_example(text):
    if pd.isna(text):
        return ""
    text = str(text)
    text = re.sub(r'[（(［\[【]\s*[✖×].*?[）)］\]】]', '', text)
    return text.strip()

def clean_japanese_example(text):
    if pd.isna(text):
        return ""
    return str(text).strip()

def split_by_markers(text, is_english=True):
    markers = r'[①②③④⑤⑥⑦⑧⑨⑩]'
    if not re.search(markers, text):
        return [line.strip() for line in text.split('\n') if line.strip()]
        
    parts = re.split(r'([①②③④⑤⑥⑦⑧⑨⑩])', text)
    sentences = []
    
    first_part = parts[0].strip()
    if first_part:
        sentences.append(first_part)
        
    for i in range(1, len(parts), 2):
        content = parts[i+1] if i+1 < len(parts) else ""
        if is_english:
            content = re.sub(r'\s*\n\s*', ' ', content)
        else:
            content = re.sub(r'\s*\n\s*', '', content)
        sentences.append(content.strip())
        
    return [s for s in sentences if s]

sheets_to_process = ['Part1', 'Part2', 'Part3', 'Part4', 'Part5', '＋α', '外来語']

output_dir = "generated_apps"
os.makedirs(output_dir, exist_ok=True)

apps_info = []

for sheet in sheets_to_process:
    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    if 'Week' not in df.columns:
        continue
    df['Week'] = df['Week'].ffill()
    
    weeks = df['Week'].unique()
    for week in weeks:
        if pd.isna(week):
            continue
        week_df = df[df['Week'] == week]
        chapter_data = []
        for idx, row in week_df.iterrows():
            word = str(row['単語']) if pd.notna(row['単語']) else ""
            explanation = str(row['語の意味']) if pd.notna(row['語の意味']) else ""
            midashi = str(row['見出番号']) if pd.notna(row['見出番号']) else str(idx)
            en_text = row['用例（英語）'] if '用例（英語）' in row else ""
            ja_text = row['用例（日本語）'] if '用例（日本語）' in row else ""
            
            if sheet in ['＋α', '外来語']:
                en_text = ""
                ja_text = ""
            
            en_clean = clean_english_example(en_text)
            ja_clean = clean_japanese_example(ja_text)
            en_lines = split_by_markers(en_clean, is_english=True)
            ja_lines = split_by_markers(ja_clean, is_english=False)
            
            if len(en_lines) > 1 and len(en_lines) == len(ja_lines):
                for i, (en_l, ja_l) in enumerate(zip(en_lines, ja_lines)):
                    en_l = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d+]\s*', '', en_l)
                    ja_l = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d+]\s*', '', ja_l)
                    en_l = re.sub(r'[あ-んア-ン一-龥]+', '', en_l).strip()
                    chapter_data.append({"id": f"{midashi}-{i+1}", "answer": word, "explanation": explanation, "en": en_l, "ja": ja_l, "target_word": word})
            else:
                en_l = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d+]\s*', '', en_clean)
                ja_l = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩\d+]\s*', '', ja_clean)
                en_l = re.sub(r'\s*\n\s*', ' ', en_l)
                ja_l = re.sub(r'\s*\n\s*', '', ja_l)
                en_l = re.sub(r'[あ-んア-ン一-龥]+', '', en_l).strip()
                chapter_data.append({"id": str(midashi), "answer": word, "explanation": explanation, "en": en_l, "ja": ja_l, "target_word": word})
                
        if not chapter_data: continue
            
        week_num = re.sub(r'[^0-9]', '', str(week))
        part_num = re.sub(r'[^0-9]', '', sheet) if 'Part' in sheet else sheet
        
        if 'Part' in sheet:
            file_name = f"part{part_num}_week{week_num}.html"
            title = f"LEAP Basic学習アプリ (P{part_num} W{week_num})"
            subtitle = f"Part {part_num} Week {week_num}"
            part_display = f"Part {part_num}"
            week_display = f"Week {week_num}"
        else:
            file_name = f"{sheet}_week{week_num}.html"
            title = f"LEAP Basic学習アプリ ({sheet} W{week_num})"
            subtitle = f"{sheet} Week {week_num}"
            part_display = sheet
            week_display = f"Week {week_num}"
            
        html_content = part1_html.replace("<title>LEAP Basic学習アプリ (P1 W1)</title>", f"<title>{title}</title>")
        html_content = html_content.replace('<h2 id="app-subtitle" class="text-xl md:text-2xl font-bold text-white">Part 1 Week 1</h2>', f'<h2 id="app-subtitle" class="text-xl md:text-2xl font-bold text-white">{subtitle}</h2>')
        
        json_data = json.dumps(chapter_data, ensure_ascii=False, indent=4)
        with open(os.path.join(output_dir, file_name), "w", encoding="utf-8") as f:
            f.write(html_content + json_data + part2_html)

        apps_info.append({
            "title": title,
            "subtitle": subtitle,
            "file_name": file_name,
            "part": part_display,
            "week": week_display,
            "word_count": len(chapter_data)
        })

# Update index.html dynamically with the collected apps_info
index_html_path = os.path.join(output_dir, "index.html")
if os.path.exists(index_html_path):
    with open(index_html_path, "r", encoding="utf-8") as f:
        index_content = f.read()
    
    apps_info_json = json.dumps(apps_info, ensure_ascii=False)
    index_content = re.sub(
        r'const appsInfo = \[.*?\];',
        f'const appsInfo = {apps_info_json};',
        index_content,
        flags=re.DOTALL
    )
    
    with open(index_html_path, "w", encoding="utf-8") as f:
        f.write(index_content)
            
print("Success")
