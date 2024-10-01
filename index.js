let openai;
let apiKey = '';
let model = 'gpt-4o-mini';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("translateToChinese").onclick = translateToChinese;
        document.getElementById("translateToEnglish").onclick = translateToEnglish;
        document.getElementById("saveSettings").onclick = saveSettings;
        
        // 加载保存的设置
        loadSettings();
    }
});

function loadSettings() {
    apiKey = localStorage.getItem('apiKey') || '';
    model = localStorage.getItem('model') || 'gpt-4o-mini';
    document.getElementById('apiKey').value = apiKey;
    document.getElementById('model').value = model;
    initializeOpenAI();
}

function saveSettings() {
    apiKey = document.getElementById('apiKey').value;
    model = document.getElementById('model').value;
    localStorage.setItem('apiKey', apiKey);
    localStorage.setItem('model', model);
    initializeOpenAI();
}

function initializeOpenAI() {
    if (apiKey) {
        openai = new OpenAI({
            apiKey: apiKey,
            dangerouslyAllowBrowser: true
        });
    } else {
        console.error("API key is not set");
    }
}

async function translateToChinese() {
    await translate('zh');
}

async function translateToEnglish() {
    await translate('en');
}

async function translate(targetLang) {
    if (!openai) {
        console.error("OpenAI is not initialized. Please set your API key.");
        return;
    }

    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load("text");
            await context.sync();

            const text = range.text;
            if (!text) {
                console.log("No text selected");
                return;
            }

            const translatedText = await translateWithOpenAI(text, targetLang);

            range.insertText(translatedText, Word.InsertLocation.replace);
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}

async function translateWithOpenAI(text, targetLang) {
    const prompt = `Translate the following ${targetLang === 'zh' ? 'English' : 'Chinese'} text to ${targetLang === 'zh' ? 'Chinese' : 'English'}:\n\n${text}`;

    try {
        const response = await openai.createCompletion({
            model: model,
            prompt: prompt,
            max_tokens: 150,
            n: 1,
            stop: null,
            temperature: 0.5,
        });

        return response.choices[0].text.trim();
    } catch (error) {
        console.error("OpenAI API error:", error);
        return "Translation error occurred";
    }
}