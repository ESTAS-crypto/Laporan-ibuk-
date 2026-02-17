// ===== VOICE.JS v10.0 — Real-Time AI Correction (Groq) =====
// Web Speech API + Enhanced Audio + Continuous Groq AI

var VoiceInput = (function () {
    var SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    var isRecording = false;
    var activeBtn = null;
    var currentTextarea = null;
    var baseText = '';
    var recognition = null;
    var allFinalText = '';
    var lockedText = '';          // Finalized text from previous sessions
    var lastRawSessionText = '';  // Raw final text from previous session (for overlap detection)
    var stopped = false;
    var silenceCount = 0;
    var MAX_SILENCE_FOR_AI = 3; // Trigger AI after 3s silence
    var MAX_SILENCE_STOP = 10;   // Stop recording after 10s silence
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

    async function saveSettings(enabled, key) {
        useAI = enabled;
        groqKey = key;
        localStorage.setItem('voice_use_ai', enabled);
        localStorage.setItem('voice_groq_key', key);
        updateUiState();

        if (enabled && key) {
            if (key.startsWith('gsk_')) {
                showMsg('🔄 Menguji koneksi API...');
                try {
                    // Simple test call
                    var response = await fetch('https://api.groq.com/openai/v1/models', {
                        headers: { 'Authorization': 'Bearer ' + key }
                    });
                    if (response.ok) {
                        showMsg('✅ Terhubung! API Key valid.');
                    } else {
                        showMsg('❌ Gagal terhubung. Cek API Key.');
                    }
                } catch (e) {
                    showMsg('⚠️ Koneksi Error');
                }
            } else {
                showMsg('⚠️ Format API Key sepertinya salah (harus gsk_...)');
            }
        } else {
            showMsg('✅ Pengaturan disimpan!');
        }
    }

    function updateUiState() {
        var status = document.getElementById('aiStatusText');
        var inputGroup = document.getElementById('apiKeyGroup');
        var toggle = document.getElementById('aiToggle');
        var keyInput = document.getElementById('groqKey');

        if (status) status.textContent = useAI ? 'Aktif (Real-Time)' : 'Non-aktif';
        if (inputGroup) inputGroup.style.display = useAI ? 'block' : 'none';
        if (toggle) toggle.checked = useAI;
        if (keyInput) keyInput.value = groqKey;
    }

    async function correctTextWithAI(text, ta, isPartial = false) {
        // Robust check: ensure ta is valid
        if (!useAI || !groqKey || !text || text.length < 5 || !ta || !ta.classList) return;

        // Visual feedback without disabling input if partial
        isCorrecting = true;
        if (!isPartial) ta.disabled = true;

        var originalBtnText = activeBtn ? activeBtn.querySelector('.voice-btn-text').textContent : '';
        if (activeBtn) activeBtn.querySelector('.voice-btn-text').textContent = '🤖';

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
                            content: 'Anda adalah asisten medis profesional Indonesia. Tugas Anda memperbaiki transkrip suara (STT) agar akurat secara medis dan tata bahasa.\n\nKAMUS ISTILAH (Gunakan ejaan ini):\n- Hipertensi, Diabetes Melitus, Kolesterol, Asam Urat, Skizofrenia\n- Paracetamol, Amoksisilin, Captopril, Metformin, Amlodipine\n- Puskesmas, Posyandu, Lansia, Balita, Ibu Hamil\n- Kunjungan Rumah, Rawat Jalan, Rujukan\n\nATURAN FORMAT:\n1. Angka Tensi: "seratus dua puluh per delapan puluh" -> "120/80 mmHg"\n2. Gula Darah: "gula darah sewaktu dua ratus" -> "GDS 200 mg/dL"\n3. Suhu: "tiga puluh enam koma lima" -> "36.5°C"\n4. Berat: "lima puluh kilo" -> "50 kg"\n\nINSTRUKSI:\n- Perbaiki typo dan ejaan yang salah.\n- Hapus kata filler (anu, hmm, apa itu).\n- JANGAN mengubah data/angka.\n- Output HANYA teks hasil perbaikan tanpa tanda kutip.'
                        },
                        {
                            role: 'user',
                            content: text
                        }
                    ],
                    temperature: 0.1
                })
            });

            var data = await response.json();
            if (data.choices && data.choices[0]) {
                var corrected = data.choices[0].message.content.trim();
                corrected = corrected.replace(/^"|"$/g, '');

                console.log('[Voice] ✨ Result:', corrected);

                // Update text
                // If partial (real-time), we replace the finalized text block
                if (isPartial) {
                    allFinalText = corrected;
                } else {
                    allFinalText = corrected;
                }

                // Force update UI
                if (currentTextarea) {
                    currentTextarea.value = baseText + allFinalText + (isRecording ? ' ' : '');
                    currentTextarea.dispatchEvent(new Event('change', { bubbles: true }));
                }

            } else {
                console.warn('[Voice] AI Error:', data);
            }
        } catch (err) {
            console.error('[Voice] Network Error:', err);
        } finally {
            if (!isPartial && ta) {
                ta.classList.remove('ai-correcting');
                ta.disabled = false;
            }
            isCorrecting = false;
            if (activeBtn) activeBtn.querySelector('.voice-btn-text').textContent = originalBtnText || '⏹️';
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

        if (isCorrecting && !stopped) return;

        cleanupRecognition();
        releaseKeepAlive();

        // ─── CRITICAL FIX: Echo & Mobile ───
        var isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);

        if (!isMobile) {
            try {
                console.log('[Voice] Acquiring mic (Desktop keep-alive)...');
                keepAliveStream = await navigator.mediaDevices.getUserMedia(AUDIO_CONSTRAINTS);

                var ctx = new AudioContext({ sampleRate: 16000 });
                var source = ctx.createMediaStreamSource(keepAliveStream);
                var compressor = ctx.createDynamicsCompressor();
                compressor.threshold.setValueAtTime(-50, ctx.currentTime);
                source.connect(compressor);

                // Fix echo 100%: Route to a dummy destination (Nowhere)
                var silentDest = ctx.createMediaStreamDestination();
                compressor.connect(silentDest);

                if (ctx.state === 'suspended') ctx.resume();
            } catch (err) {
                console.warn('[Voice] Mic keep-alive failed (non-fatal):', err);
            }
        } else {
            console.log('[Voice] Mobile detected: Skipping keep-alive audio pipeline');
        }

        await delay(400);

        currentTextarea = ta;
        baseText = ta.value || '';
        if (baseText && !baseText.endsWith(' ')) baseText += ' ';
        allFinalText = '';
        lockedText = '';
        lastRawSessionText = '';
        silenceCount = 0;
        stopped = false;
        isRecording = true;
        activeBtn = btn;
        btn.classList.add('recording');
        btn.querySelector('.voice-btn-text').textContent = '⏹️';

        beginListening();
        showMsg('🎙️ Bicara sekarang...' + (useAI ? ' (Real-Time Mode)' : ''));
    }

    function beginListening() {
        if (stopped || !currentTextarea) return;

        recognition = new SpeechRecognition();
        recognition.lang = 'id-ID';
        recognition.continuous = true;
        recognition.interimResults = true;
        recognition.maxAlternatives = 3;

        // This session's raw final text (rebuilt from event.results each time)
        var currentRawFinal = '';

        recognition.onresult = function (event) {
            silenceCount = 0;

            // Rebuild ALL finals and interims from event.results
            var rawFinal = '';
            var interimParts = '';
            for (var i = 0; i < event.results.length; i++) {
                var best = pickBestAlternative(event.results[i]);
                if (event.results[i].isFinal) {
                    rawFinal += best + ' ';
                } else {
                    interimParts += best;
                }
            }

            currentRawFinal = rawFinal;

            // Mobile overlap detection (Robust Word-Based)
            var contribution = rawFinal;
            if (lastRawSessionText) {
                contribution = stripOverlap(lastRawSessionText, rawFinal);
            }

            // Rebuild: lockedText (previous sessions) + this session's contribution
            allFinalText = lockedText + (contribution.trim() ? postProcess(contribution) + ' ' : '');

            // Helper function for overlap
            function stripOverlap(prev, curr) {
                var p = prev.trim();
                var c = curr.trim();
                if (!p || !c) return curr;

                // 1. Direct Prefix match
                if (c.toLowerCase().startsWith(p.toLowerCase())) {
                    return curr.substring(prev.length).trim();
                }

                // 2. Word-based Suffix/Prefix match (for partial overlaps)
                var pWords = p.toLowerCase().split(/\s+/);
                var cWords = c.toLowerCase().split(/\s+/);
                var maxWords = Math.min(pWords.length, cWords.length, 10); // Check last 10 words

                for (var n = maxWords; n > 0; n--) {
                    var suffix = pWords.slice(-n).join(' ');
                    var prefix = cWords.slice(0, n).join(' ');
                    if (suffix === prefix) {
                        // Found overlap of n words. 
                        // Return the substring of curr after these n words.
                        // We use the original 'curr' string to preserve casing/spacing of the new part.
                        // Rough logic: find index of Nth space? 
                        // Better: split original curr and join rest.
                        var words = curr.trim().split(/\s+/);
                        if (words.length > n) {
                            return words.slice(n).join(' ');
                        } else {
                            return ''; // Full overlap
                        }
                    }
                }
                return curr;
            }

            if (currentTextarea) {
                currentTextarea.value = baseText + allFinalText + interimParts;
                currentTextarea.scrollTop = currentTextarea.scrollHeight;
            }
        };

        recognition.onerror = function (event) {
            if (event.error === 'no-speech' || event.error === 'aborted') return;
        };

        recognition.onend = function () {
            // Lock in this session's contribution
            lockedText = allFinalText;
            lastRawSessionText = currentRawFinal;

            if (stopped) {
                finalize();
                return;
            }
            silenceCount++;

            // Continuous AI Trigger
            if (useAI && silenceCount >= MAX_SILENCE_FOR_AI && allFinalText.length > 5 && !isCorrecting) {
                console.log('[Voice] Silence detected, triggering partial AI correction...');
                correctTextWithAI(allFinalText, currentTextarea, true);
            }

            if (silenceCount > MAX_SILENCE_STOP) {
                showMsg('⏸️ Hening (Stop).');
                fullStop();
                return;
            }

            // Restart listener
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
            var ta = currentTextarea;
            var finalText = postProcess(allFinalText);

            ta.value = baseText + finalText;
            ta.dispatchEvent(new Event('change', { bubbles: true }));

            fullStop();

            // Final AI Correction (if not already correcting or just to be sure)
            if (useAI && finalText.trim().length > 5 && ta && !isCorrecting) {
                correctTextWithAI(finalText, ta, false);
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
        lockedText = '';
        lastRawSessionText = '';
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
            document.getElementById('aiStatusText').textContent = isChecked ? 'Aktif (Real-Time)' : 'Non-aktif';
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
    console.log('[Voice] v10.0 Real-Time Ready');
});
