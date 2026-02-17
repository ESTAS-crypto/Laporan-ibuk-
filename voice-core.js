// ===== VOICE.JS v9.3 — Hybrid AI Correction (Groq) =====
// Web Speech API + Enhanced Audio + Groq AI Post-processing

var VoiceInput = (function () {
    var SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    var isRecording = false;
    var activeBtn = null;
    var currentTextarea = null;
    var baseText = '';
    var recognition = null;
    var allFinalText = '';
    var stopped = false;
    var silenceCount = 0;
    var MAX_SILENCE = 8;
    var keepAliveStream = null;

    // AI Correction State
    var useAI = false;
    var groqKey = '';
    var isCorrecting = false;

    var isSupported = !!SpeechRecognition;

    // ─── Settings & AI Logic ───
    function loadSettings() {
        useAI = localStorage.getItem('voice_use_ai') === 'true';
        groqKey = localStorage.getItem('voice_groq_key') || '';
        updateUiState();
    }

    function saveSettings(enabled, key) {
        useAI = enabled;
        groqKey = key;
        localStorage.setItem('voice_use_ai', enabled);
        localStorage.setItem('voice_groq_key', key);
        updateUiState();
        showMsg('✅ Pengaturan disimpan!');
    }

    function updateUiState() {
        var status = document.getElementById('aiStatusText');
        var inputGroup = document.getElementById('apiKeyGroup');
        var toggle = document.getElementById('aiToggle');
        var keyInput = document.getElementById('groqKey');

        if (status) status.textContent = useAI ? 'Aktif (Hybrid Mode)' : 'Non-aktif';
        if (inputGroup) inputGroup.style.display = useAI ? 'block' : 'none';
        if (toggle) toggle.checked = useAI;
        if (keyInput) keyInput.value = groqKey;
    }

    async function correctTextWithAI(text, ta) {
        // Robust check: ensure ta is valid and still in document or at least has classList
        if (!useAI || !groqKey || !text || text.length < 5 || !ta || !ta.classList) return;

        // Visual feedback
        try {
            ta.classList.add('ai-correcting');
            ta.disabled = true;
        } catch (e) {
            console.warn('[Voice] UI update failed:', e);
            return;
        }
        isCorrecting = true;
        console.log('[Voice] 🤖 AI Correcting:', text);

        try {
            var response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
                method: 'POST',
                headers: {
                    'Authorization': 'Bearer ' + groqKey,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    model: 'llama-3.3-70b-versatile',
                    messages: [
                        {
                            role: 'system',
                            content: 'Anda adalah asisten medis Indonesia. Tugas: Perbaiki tata bahasa, ejaan, dan istilah medis dari teks input. JANGAN mengubah makna. JANGAN menambah komentar. Output HANYA teks yang sudah diperbaiki.'
                        },
                        {
                            role: 'user',
                            content: text
                        }
                    ],
                    temperature: 0.2
                })
            });

            var data = await response.json();
            if (data.choices && data.choices[0]) {
                var corrected = data.choices[0].message.content.trim();
                // Remove quotes if AI added them
                corrected = corrected.replace(/^"|"$/g, '');

                console.log('[Voice] ✨ Result:', corrected);

                // Update textarea with corrected text
                // Note: currentTextarea is null here, but ta is the valid element we captured
                ta.value = baseText + corrected;
                allFinalText = corrected; // Update memory
            } else {
                console.warn('[Voice] AI Error:', data);
                showMsg('⚠️ AI Error: ' + (data.error?.message || 'Unknown'));
            }
        } catch (err) {
            console.error('[Voice] Network Error:', err);
        } finally {
            ta.classList.remove('ai-correcting');
            ta.disabled = false;
            isCorrecting = false;
            ta.dispatchEvent(new Event('change', { bubbles: true }));
        }
    }

    // ─── Audio Constraints ───
    var AUDIO_CONSTRAINTS = {
        audio: {
            channelCount: 1,
            sampleRate: { ideal: 16000 },
            echoCancellation: true,
            noiseSuppression: true,
            autoGainControl: true
        }
    };

    // ─── Local Post-Processing (Regex) ───
    var CORRECTIONS = [
        [/\btensi\b/gi, 'tensi'], [/\btensih?\b/gi, 'tensi'],
        [/\bhiper ?tensi\b/gi, 'hipertensi'], [/\bdia ?betes\b/gi, 'diabetes'],
        [/\bkolest(?:r|er)ol\b/gi, 'kolesterol'], [/\basam ?urat\b/gi, 'asam urat'],
        [/\bgula ?darah\b/gi, 'gula darah'], [/\btekanan ?darah\b/gi, 'tekanan darah'],
        [/\bimunisas\b/gi, 'imunisasi'], [/\bkonsulas\b/gi, 'konsultasi'],
        [/\s+/g, ' '], [/\s*,\s*/g, ', '], [/\s*\.\s*/g, '. '],
        [/^\s+/, ''], [/\s+$/, '']
    ];

    function postProcess(text) {
        if (!text) return text;
        var result = text;
        for (var i = 0; i < CORRECTIONS.length; i++) {
            result = result.replace(CORRECTIONS[i][0], CORRECTIONS[i][1]);
        }
        result = result.replace(/(^|\.\s+)([a-z])/g, function (m, sep, ch) {
            return sep + ch.toUpperCase();
        });
        return result.trim();
    }

    // ─── Start Voice ───
    async function startRecording(textareaId, btn) {
        var ta = document.getElementById(textareaId);
        if (!ta) return;

        // Check Connection (PWA Requirement)
        if (!navigator.onLine) {
            showMsg('🌐 Voice butuh internet.');
            return;
        }

        if (isCorrecting) return; // Wait for AI

        cleanupRecognition();
        releaseKeepAlive();

        // Detect Mobile (Android/iOS)
        var isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);

        // Skip audio pipeline on mobile (prevent echo & mic conflict)
        if (!isMobile) {
            try {
                console.log('[Voice] Acquiring mic (Desktop keep-alive)...');
                keepAliveStream = await navigator.mediaDevices.getUserMedia(AUDIO_CONSTRAINTS);

                // Activate Audio Pipeline
                var ctx = new AudioContext({ sampleRate: 16000 });
                var source = ctx.createMediaStreamSource(keepAliveStream);
                var compressor = ctx.createDynamicsCompressor();
                compressor.threshold.setValueAtTime(-50, ctx.currentTime);
                source.connect(compressor);

                // Fix echo 100%: Route to a dummy destination (Nowhere)
                // Do NOT connect to ctx.destination (Speakers)
                var silentDest = ctx.createMediaStreamDestination();
                compressor.connect(silentDest);

                // Keep context running
                if (ctx.state === 'suspended') ctx.resume();
            } catch (err) {
                console.warn('[Voice] Mic keep-alive failed (non-fatal):', err);
                // Continue anyway - SpeechRecognition might still work
            }
        } else {
            console.log('[Voice] Mobile detected: Skipping keep-alive audio pipeline');
        }

        await delay(400);

        currentTextarea = ta;
        baseText = ta.value || '';
        if (baseText && !baseText.endsWith(' ')) baseText += ' ';
        allFinalText = '';
        recentFinals = [];
        silenceCount = 0;
        stopped = false;
        isRecording = true;
        activeBtn = btn;
        btn.classList.add('recording');
        btn.querySelector('.voice-btn-text').textContent = '⏹️';

        beginListening();
        showMsg('🎙️ Bicara sekarang...' + (useAI ? ' (Hybrid Mode ON)' : ''));
    }

    // Deduplication: track recent final texts to prevent mobile repeat
    var recentFinals = [];
    var MAX_RECENT = 5;

    function isDuplicate(text) {
        var clean = text.trim().toLowerCase();
        if (!clean) return true;
        for (var i = 0; i < recentFinals.length; i++) {
            var recent = recentFinals[i];
            if (clean === recent) return true;
            if (recent.indexOf(clean) !== -1) return true;
            if (clean.indexOf(recent) !== -1 && clean.length - recent.length < 5) return true;
        }
        return false;
    }

    function addToRecent(text) {
        var clean = text.trim().toLowerCase();
        if (!clean) return;
        recentFinals.push(clean);
        if (recentFinals.length > MAX_RECENT) recentFinals.shift();
    }

    function beginListening() {
        if (stopped || !currentTextarea) return;

        recognition = new SpeechRecognition();
        recognition.lang = 'id-ID';
        recognition.continuous = true;
        recognition.interimResults = true;
        recognition.maxAlternatives = 3;

        // Track how many results from this session we've already finalized
        var processedFinalCount = 0;

        recognition.onresult = function (event) {
            silenceCount = 0;
            var newFinalText = '';
            var interimParts = '';

            for (var i = 0; i < event.results.length; i++) {
                var best = pickBestAlternative(event.results[i]);
                if (event.results[i].isFinal) {
                    // Only process results we haven't already added
                    if (i >= processedFinalCount) {
                        // Mobile deduplication: skip if this text was recently finalized
                        if (!isDuplicate(best)) {
                            newFinalText += best + ' ';
                            addToRecent(best);
                        } else {
                            console.log('[Voice] Skipped duplicate:', best);
                        }
                        processedFinalCount = i + 1;
                    }
                } else {
                    interimParts += best;
                }
            }

            // Append only NEW final text
            if (newFinalText) {
                allFinalText += postProcess(newFinalText);
            }

            if (currentTextarea) {
                currentTextarea.value = baseText + allFinalText + interimParts;
                currentTextarea.style.height = 'auto';
                currentTextarea.style.height = currentTextarea.scrollHeight + 'px';
            }
        };

        recognition.onerror = function (event) {
            if (event.error === 'no-speech' || event.error === 'aborted') return;
            showMsg('⚠️ Error: ' + event.error);
            fullStop();
        };

        recognition.onend = function () {
            if (stopped) {
                finalize();
                return;
            }
            silenceCount++;
            if (silenceCount > MAX_SILENCE) {
                showMsg('⏸️ Hening.');
                fullStop();
                return;
            }
            // Restart listener (new session — processedFinalCount resets via new closure)
            setTimeout(function () {
                if (!stopped) beginListening();
            }, 300);
        };

        try { recognition.start(); } catch (e) { fullStop(); }
    }

    function pickBestAlternative(result) {
        if (result.length <= 1) return result[0].transcript;
        var best = result[0];
        for (var i = 1; i < result.length; i++) {
            if (result[i].confidence > best.confidence) best = result[i];
        }
        return best.transcript;
    }

    function stopRecording() {
        stopped = true;
        if (recognition) try { recognition.stop(); } catch (e) { }
        setTimeout(function () { if (activeBtn) finalize(); }, 1500);
    }

    function finalize() {
        if (currentTextarea) {
            var ta = currentTextarea; // Capture reference
            var finalText = postProcess(allFinalText);

            ta.value = baseText + finalText;
            ta.dispatchEvent(new Event('change', { bubbles: true }));

            fullStop(); // Resets global currentTextarea to null

            // Trigger AI Correction if enabled
            if (useAI && finalText.trim().length > 5 && ta) {
                correctTextWithAI(finalText, ta);
            }
        } else {
            fullStop();
        }
    }

    function cleanupRecognition() {
        if (recognition) { try { recognition.abort(); } catch (e) { } recognition = null; }
    }

    function releaseKeepAlive() {
        if (keepAliveStream) { keepAliveStream.getTracks().forEach(function (t) { t.stop(); }); keepAliveStream = null; }
    }

    function fullStop() {
        stopped = true;
        isRecording = false;
        cleanupRecognition();
        releaseKeepAlive();
        if (activeBtn) {
            activeBtn.classList.remove('recording');
            activeBtn.querySelector('.voice-btn-text').textContent = '🎤';
            activeBtn = null;
        }
        currentTextarea = null;
        allFinalText = '';
        silenceCount = 0;
    }

    function delay(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

    function showMsg(text) {
        var existing = document.querySelectorAll('.voice-alert');
        for (var i = 0; i < existing.length; i++) existing[i].remove();
        var el = document.createElement('div');
        el.className = 'voice-alert';
        el.textContent = text;
        document.body.appendChild(el);
        setTimeout(function () { el.classList.add('show'); }, 10);
        setTimeout(function () {
            el.classList.remove('show');
            setTimeout(function () { el.remove(); }, 300);
        }, 4000);
    }

    function createButton(textareaId) {
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn-voice';
        btn.innerHTML = '<span class="voice-btn-text">🎤</span>';
        if (!isSupported) { btn.disabled = true; return btn; }

        btn.addEventListener('click', function (e) {
            e.preventDefault();
            e.stopPropagation();
            if (isRecording && activeBtn === btn) stopRecording();
            else if (isRecording) {
                fullStop();
                setTimeout(function () { startRecording(textareaId, btn); }, 300);
            } else {
                startRecording(textareaId, btn);
            }
        });
        return btn;
    }

    function initForLaporan(idx) {
        var kl = document.querySelector('label[for="keluhan_' + idx + '"]');
        var tl = document.querySelector('label[for="tindakLanjut_' + idx + '"]');
        if (kl && !kl.querySelector('.btn-voice')) kl.appendChild(createButton('keluhan_' + idx));
        if (tl && !tl.querySelector('.btn-voice')) tl.appendChild(createButton('tindakLanjut_' + idx));
    }

    // ─── Settings Modal Logic ───
    function initSettings() {
        var modal = document.getElementById('settingsModal');
        var btn = document.getElementById('btnSettings');
        var close = document.querySelector('.close-modal');
        var save = document.getElementById('saveSettings');
        var toggle = document.getElementById('aiToggle');

        if (!modal || !btn) return;

        btn.onclick = function () { modal.style.display = 'block'; loadSettings(); };
        close.onclick = function () { modal.style.display = 'none'; };
        window.onclick = function (e) { if (e.target == modal) modal.style.display = 'none'; };

        toggle.onchange = function () {
            var isChecked = this.checked;
            document.getElementById('apiKeyGroup').style.display = isChecked ? 'block' : 'none';
            document.getElementById('aiStatusText').textContent = isChecked ? 'Aktif (Hybrid Mode)' : 'Non-aktif';
        };

        save.onclick = function () {
            var key = document.getElementById('groqKey').value.trim();
            saveSettings(toggle.checked, key);
            modal.style.display = 'none';
        };

        loadSettings();
    }

    return {
        initForLaporan: initForLaporan,
        initSettings: initSettings,
        stop: stopRecording
    };
})();

document.addEventListener('DOMContentLoaded', function () {
    VoiceInput.initForLaporan(1);
    VoiceInput.initSettings();
    console.log('[Voice] v9.3 Hybrid Ready');
});
