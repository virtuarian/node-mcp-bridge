import { t, getCurrentLanguage, setLanguage } from './i18n.js';

document.addEventListener('DOMContentLoaded', () => {
    // 言語セレクターの初期化
    const languageSelector = document.getElementById('languageSelector');
    languageSelector.value = getCurrentLanguage();
    languageSelector.addEventListener('change', (e) => {
        setLanguage(e.target.value);
    });

    // 画面上のすべてのテキストを初期化
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        el.textContent = t(key);
    });
    
    // プレースホルダーテキストの多言語対応
    document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
        const key = el.getAttribute('data-i18n-placeholder');
        el.placeholder = t(key);
    });
    
    // タイトル属性の多言語対応
    document.querySelectorAll('[data-i18n-title]').forEach(el => {
        const key = el.getAttribute('data-i18n-title');
        el.title = t(key);
    });

    // モーダルタイトルの多言語化
    document.querySelectorAll('.modal-title:not([data-i18n])').forEach(el => {
        // IDに基づいてタイトルを設定
        if (el.id === 'confirmMessage' || el.parentElement.parentElement.id === 'confirmModal') {
            el.textContent = t('confirm.title');
        }
    });

    // 要素の取得
    const serverList = document.getElementById('serverList');
    const refreshBtn = document.getElementById('refreshBtn');
    const addServerBtn = document.getElementById('addServerBtn');
    const serverModal = document.getElementById('serverModal');
    const serverModalTitle = document.getElementById('serverModalTitle');
    const closeServerModal = document.getElementById('closeServerModal');
    const serverForm = document.getElementById('serverForm');
    const serverAction = document.getElementById('serverAction');
    const serverOriginalName = document.getElementById('serverOriginalName');
    const serverName = document.getElementById('serverName');
    const serverCommand = document.getElementById('serverCommand');
    const serverArgs = document.getElementById('serverArgs');
    const envsBody = document.getElementById('envsBody');
    const addEnvBtn = document.getElementById('addEnvBtn');
    const serverTimeout = document.getElementById('serverTimeout');
    const serverSessionTimeout = document.getElementById('serverSessionTimeout');
    const serverDisabled = document.getElementById('serverDisabled');
    const cancelServerBtn = document.getElementById('cancelServerBtn');
    const saveServerBtn = document.getElementById('saveServerBtn');
    const alertContainer = document.getElementById('alertContainer');
    const confirmModal = document.getElementById('confirmModal');
    const confirmMessage = document.getElementById('confirmMessage');
    const closeConfirmModal = document.getElementById('closeConfirmModal');
    const cancelConfirmBtn = document.getElementById('cancelConfirmBtn');
    const confirmBtn = document.getElementById('confirmBtn');
    const serverDetailsModal = document.getElementById('serverDetailsModal');
    const closeServerDetailsModal = document.getElementById('closeServerDetailsModal');
    const closeServerDetailsBtn = document.getElementById('closeServerDetailsBtn');

    let confirmCallback = null;


    closeServerDetailsModal.addEventListener('click', () => {
        closeModal(serverDetailsModal);
    });

    closeServerDetailsBtn.addEventListener('click', () => {
        closeModal(serverDetailsModal);
    });


    // サーバー一覧取得・表示部分を修正
    async function fetchAndRenderServers() {
        try {
            const response = await fetch('/admin/servers');
            const servers = await response.json();
            const serverCards = document.getElementById('serverCards');

            if (servers.length === 0) {
                serverCards.innerHTML = `<div class="empty-state">${t('empty.servers')}</div>`;
                return;
            }

            serverCards.innerHTML = '';

            servers.forEach(server => {
                // Parse config to get command and timeout
                const config = JSON.parse(server.config);

                // Status class
                let statusClass = '';
                let statusText = '';
                
                switch (server.status) {
                    case 'connected':
                        statusClass = 'status-connected';
                        statusText = t('server.status.connected');
                        break;
                    case 'connecting':
                        statusClass = 'status-connecting';
                        statusText = t('server.status.connecting');
                        break;
                    case 'disconnected':
                        statusClass = 'status-disconnected';
                        statusText = t('server.status.disconnected');
                        break;
                }

                const errorMessage = server.error 
                    ? `<div class="error-message">${t('server.error', server.error.split('\n')[0])}</div>` 
                    : '';

                const card = document.createElement('div');
                card.className = 'server-card';
                card.setAttribute('data-name', server.name);

                card.innerHTML = `
                    <div class="server-card-header">
                        <h3 class="server-name">${server.name}</h3>
                        <span class="server-status ${statusClass}">
                            <span class="status-dot"></span>
                            ${statusText}
                        </span>
                    </div>
                    ${errorMessage}
                    <div class="server-card-content">
                        <div class="server-info-grid">
                            <div class="server-info-item">
                                <div class="info-label">${t('server.command')}:</div>
                                <div class="info-value">${config.command} ${config.args ? config.args.join(' ') : ''}</div>
                            </div>
                            <div class="server-info-item">
                                <div class="info-label">${t('server.timeout')}:</div>
                                <div class="info-value">${config.timeout || '120'} ${t('server.seconds')}</div>
                            </div>
                            <div class="server-info-item">
                                <div class="info-label">${t('server.enable')}:</div>
                                <div class="info-value">
                                    <label class="switch">
                                        <input type="checkbox" class="server-toggle" ${!server.disabled ? 'checked' : ''}>
                                        <span class="slider"></span>
                                    </label>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="server-card-actions">
                        <button class="btn btn-icon btn-outline details-btn" title="${t('server.details')}">
                            <span class="material-icons">info</span>
                            <span class="sr-only">${t('server.details')}</span>
                        </button>
                        <button class="btn btn-icon btn-outline restart-btn" title="${t('server.restart')}">
                            <span class="material-icons">refresh</span>
                            <span class="sr-only">${t('server.restart')}</span>
                        </button>
                        <button class="btn btn-icon btn-outline edit-btn" title="${t('server.edit')}">
                            <span class="material-icons">edit</span>
                            <span class="sr-only">${t('server.edit')}</span>
                        </button>
                        <button class="btn btn-icon btn-danger delete-btn" title="${t('server.delete')}">
                            <span class="material-icons">delete</span>
                            <span class="sr-only">${t('server.delete')}</span>
                        </button>
                    </div>
                `;

                // イベントリスナーを追加（既存のコードと同様）
                // 詳細ボタン
                const detailsBtn = card.querySelector('.details-btn');
                detailsBtn.addEventListener('click', () => {
                    showServerDetails(server);
                });

                // Toggle server enabled/disabled
                const toggle = card.querySelector('.server-toggle');
                toggle.addEventListener('change', async () => {
                    await toggleServerStatus(server.name, !toggle.checked);
                });

                // Restart button
                const restartBtn = card.querySelector('.restart-btn');
                restartBtn.addEventListener('click', async () => {
                    await restartServer(server.name);
                });

                // Edit button
                const editBtn = card.querySelector('.edit-btn');
                editBtn.addEventListener('click', () => {
                    editServer(server);
                });

                // Delete button
                const deleteBtn = card.querySelector('.delete-btn');
                deleteBtn.addEventListener('click', () => {
                    showConfirmation(
                        t('confirm.delete', server.name),
                        async () => {
                            await deleteServer(server.name);
                        }
                    );
                });

                serverCards.appendChild(card);
            });
        } catch (error) {
            const serverCards = document.getElementById('serverCards');
            serverCards.innerHTML = `<div class="error-message">${t('error.serverList', error.message)}</div>`;
        }
    }

    // サーバーステータスのトグル
    async function toggleServerStatus(serverName, disabled) {
        try {
            const response = await fetch(`/admin/servers/${serverName}/toggleDisabled`, {
                method: 'PUT',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ disabled })
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || 'サーバーステータスの更新に失敗しました');
            }

            showAlert(
                'success',
                t('alert.toggle.success',
                    serverName,
                    disabled ? t('alert.toggle.disabled') : t('alert.toggle.enabled'))
            );

            await fetchAndRenderServers();
        } catch (error) {
            showAlert('danger', error.message);
            // 失敗したらUIを元に戻す
            await fetchAndRenderServers();
        }
    }

    // サーバーの再起動
    async function restartServer(serverName) {
        try {
            // セレクタを修正してカード内のボタンを正しく取得
            const btn = document.querySelector(`.server-card[data-name="${serverName}"] .restart-btn`);
            const originalText = btn.innerHTML;
            
            btn.innerHTML = `<span class="loading"></span>${t('loading')}`;
            btn.disabled = true;

            const response = await fetch(`/admin/servers/${serverName}/restart`, {
                method: 'POST'
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || t('error.serverRestart', ''));
            }

            showAlert('success', t('alert.restart.success', serverName));

            // 再起動後に状態を更新するために少し待つ
            setTimeout(async () => {
                await fetchAndRenderServers();
                btn.innerHTML = originalText;
                btn.disabled = false;
            }, 2000);
        } catch (error) {
            showAlert('danger', error.message);
            await fetchAndRenderServers();
        }
    }

    // サーバーの削除
    async function deleteServer(serverName) {
        try {
            const response = await fetch(`/admin/servers/${serverName}`, {
                method: 'DELETE'
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || t('error.serverDelete', ''));
            }

            showAlert('success', t('alert.delete.success', serverName));
            await fetchAndRenderServers();
        } catch (error) {
            showAlert('danger', error.message);
        }
    }

    // サーバーの編集関数を更新
    function editServer(server) {
        const config = JSON.parse(server.config);

        serverAction.value = 'edit';
        serverOriginalName.value = server.name;
        serverModalTitle.textContent = t('server.edit');

        serverName.value = server.name;
        serverCommand.value = config.command || '';
        serverCommand.placeholder = t('server.command.placeholder');
        serverArgs.value = config.args ? config.args.join('\n') : '';
        serverArgs.placeholder = t('server.args.placeholder');
        serverTimeout.value = config.timeout || 120;
        serverSessionTimeout.value = config.sessionTimeout !== undefined ? config.sessionTimeout : 180;
        serverDisabled.checked = !!server.disabled;

        // 環境変数のクリアと設定
        envsBody.innerHTML = '';
        if (config.env) {
            Object.entries(config.env).forEach(([key, value]) => {
                addEnvRow(key, value);
            });
        }

        // ヒントテキストの多言語対応
        document.querySelector('.form-hint').textContent = t('server.sessionTimeout.hint');

        openModal(serverModal);
    }

    // 環境変数の行を追加する関数も多言語対応

    function addEnvRow(key = '', value = '') {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><input type="text" class="form-control env-key" value="${key}"></td>
            <td><input type="text" class="form-control env-value" value="${value}"></td>
            <td>
                <button type="button" class="btn btn-icon btn-sm btn-danger remove-env" title="${t('env.delete')}">
                    <span class="material-icons" style="font-size: 16px;">delete</span>
                    <span class="sr-only">${t('env.delete')}</span>
                </button>
            </td>
        `;

        const removeBtn = row.querySelector('.remove-env');
        removeBtn.addEventListener('click', () => {
            row.remove();
        });

        envsBody.appendChild(row);
    }

    // 新しいサーバーの追加
    function resetServerForm() {
        serverAction.value = 'add';
        serverOriginalName.value = '';
        serverModalTitle.textContent = t('server.add');

        serverForm.reset();
        serverCommand.placeholder = t('server.command.placeholder');
        serverArgs.placeholder = t('server.args.placeholder');
        document.querySelector('.form-hint').textContent = t('server.sessionTimeout.hint');
        
        envsBody.innerHTML = '';
        addEnvRow(); // デフォルトで1行追加
    }

    // サーバーの保存関数を更新
    async function saveServer() {
        const name = serverName.value.trim();
        const command = serverCommand.value.trim();
        const args = serverArgs.value
            ? serverArgs.value.split('\n').map(arg => arg.trim()).filter(arg => arg)
            : [];
        const timeout = parseInt(serverTimeout.value, 10) || 120;
        const sessionTimeout = serverSessionTimeout.value.trim() === ''
            ? undefined
            : parseInt(serverSessionTimeout.value, 10);
        const disabled = serverDisabled.checked;

        // 環境変数の取得
        const env = {};
        document.querySelectorAll('#envsBody tr').forEach(row => {
            const keyInput = row.querySelector('.env-key');
            const valueInput = row.querySelector('.env-value');

            if (keyInput && valueInput) {
                const key = keyInput.value.trim();
                const value = valueInput.value.trim();

                if (key) {
                    env[key] = value;
                }
            }
        });

        // サーバー設定オブジェクトの作成
        const serverConfig = {
            command,
            args,
            env,
            timeout,
            sessionTimeout,
            disabled,
            autoApprove: [] // デフォルトで空の配列
        };

        try {
            const isEdit = serverAction.value === 'edit';
            const url = isEdit ? `/admin/servers/${serverOriginalName.value}` : `/admin/servers/${name}`;

            const response = await fetch(url, {
                method: 'PUT',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(serverConfig)
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || t('error.serverUpdate', ''));
            }

            showAlert(
                'success',
                t('alert.update.success', name, t(isEdit ? 'alert.update.updated' : 'alert.update.added'))
            );
            closeModal(serverModal);
            await fetchAndRenderServers();
        } catch (error) {
            showAlert('danger', error.message);
        }
    }

    // モーダルを開く
    function openModal(modal) {
        modal.classList.add('active');
    }

    // モーダルを閉じる
    function closeModal(modal) {
        modal.classList.remove('active');
    }

    // 確認ダイアログを表示
    function showConfirmation(message, callback) {
        confirmMessage.textContent = message;
        confirmCallback = callback;
        openModal(confirmModal);
    }

    // アラートの表示
    function showAlert(type, message) {
        const alert = document.createElement('div');
        alert.className = `alert alert-${type}`;
        alert.textContent = message;

        alertContainer.innerHTML = '';
        alertContainer.appendChild(alert);

        // 5秒後に自動的に消える
        setTimeout(() => {
            alert.remove();
        }, 5000);
    }

    // イベントリスナーの設定
    refreshBtn.addEventListener('click', fetchAndRenderServers);

    addServerBtn.addEventListener('click', () => {
        resetServerForm();
        openModal(serverModal);
    });

    closeServerModal.addEventListener('click', () => {
        closeModal(serverModal);
    });

    cancelServerBtn.addEventListener('click', () => {
        closeModal(serverModal);
    });

    saveServerBtn.addEventListener('click', saveServer);

    addEnvBtn.addEventListener('click', () => {
        addEnvRow();
    });

    closeConfirmModal.addEventListener('click', () => {
        closeModal(confirmModal);
    });

    cancelConfirmBtn.addEventListener('click', () => {
        closeModal(confirmModal);
    });

    confirmBtn.addEventListener('click', async () => {
        if (confirmCallback) {
            await confirmCallback();
            confirmCallback = null;
        }
        closeModal(confirmModal);
    });

    // モーダルオーバーレイのクリックで閉じる機能を追加
    const modalOverlays = document.querySelectorAll('.modal-overlay');
    modalOverlays.forEach(overlay => {
        overlay.addEventListener('click', (event) => {
            // オーバーレイ自体がクリックされた場合のみ閉じる
            if (event.target === overlay) {
                closeModal(overlay);
            }
        });
    });

    // モーダル本体へのクリック伝播を防止
    const modals = document.querySelectorAll('.modal');
    modals.forEach(modal => {
        modal.addEventListener('click', (event) => {
            // モーダル内のクリックがオーバーレイに伝播しないようにする
            event.stopPropagation();
        });
    });

    // サーバー詳細を表示
    async function showServerDetails(server) {
        // DOM要素の取得
        const serverDetailsModal = document.getElementById('serverDetailsModal');
        const serverDetailsTitle = document.getElementById('serverDetailsTitle');
        const serverBasicInfo = document.getElementById('serverBasicInfo');
        const serverToolsList = document.getElementById('serverToolsList');
        const toolsLoadingSpinner = document.getElementById('toolsLoadingSpinner');
        const serverEnvVars = document.getElementById('serverEnvVars');

        // タイトル設定
        serverDetailsTitle.textContent = t('server.details.title', server.name);

        // 基本情報の表示
        const config = JSON.parse(server.config);

        // セッションタイムアウト表示を追加
        let sessionTimeoutDisplay = config.sessionTimeout === 0
            ? t('server.timeout.unlimited')
            : (config.sessionTimeout || 180) + ' ' + t('minutes');

        serverBasicInfo.innerHTML = `
            <div class="info-row">
                <div class="info-label">${t('server.label.name')}</div>
                <div class="info-value">${server.name}</div>
            </div>
            <div class="info-row">
                <div class="info-label">${t('server.label.status')}</div>
                <div class="info-value">
                    <span class="server-status status-${server.status}">
                        <span class="status-dot"></span>
                        ${t('server.status.' + server.status)}
                    </span>
                </div>
            </div>
            <div class="info-row">
                <div class="info-label">${t('server.label.command')}</div>
                <div class="info-value">${config.command}</div>
            </div>
            <div class="info-row">
                <div class="info-label">${t('server.label.args')}</div>
                <div class="info-value">${config.args ? config.args.join('<br>') : t('server.args.none')}</div>
            </div>
            <div class="info-row">
                <div class="info-label">${t('server.sessionTimeout')}</div>
                <div class="info-value">${sessionTimeoutDisplay}</div>
            </div>
            <div class="info-row">
                <div class="info-label">${t('server.label.timeout')}</div>
                <div class="info-value">${config.timeout || '120'} ${t('server.seconds')}</div>
            </div>
            <div class="info-row">
                <div class="info-label">${t('server.label.state')}</div>
                <div class="info-value">${server.disabled ? t('server.disabled') : t('server.enabled')}</div>
            </div>
        `;

        // サーバー詳細表示関数内の修正

        // モーダルを表示する前に詳細画面のタイトルもi18n化
        const detailSectionTitles = serverDetailsModal.querySelectorAll('.server-info-section h4');
        detailSectionTitles.forEach(title => {
            const i18nKey = title.getAttribute('data-i18n');
            if (i18nKey) {
                title.textContent = t(i18nKey);
            }
        });

        // 環境変数のテーブルヘッダーも多言語化
        if (config.env && Object.keys(config.env).length > 0) {
            let envHtml = `<table class="details-table"><thead><tr><th>${t('server.env.varName')}</th><th>${t('server.env.value')}</th></tr></thead><tbody>`;

            Object.entries(config.env).forEach(([key, value]) => {
                envHtml += `<tr><td>${key}</td><td>${value}</td></tr>`;
            });

            envHtml += '</tbody></table>';
            serverEnvVars.innerHTML = envHtml;
        } else {
            serverEnvVars.innerHTML = `<p>${t('server.details.noEnvVars')}</p>`;
        }

        // モーダルを表示
        openModal(serverDetailsModal);

        // ツール一覧のリセットと読み込み中表示
        serverToolsList.innerHTML = '';
        toolsLoadingSpinner.style.display = 'flex';

        try {
            // ツール一覧の取得と表示
            const tools = await fetchServerTools(server.name);
            toolsLoadingSpinner.style.display = 'none';

            // showServerDetails 関数内のツール表示部分を修正
            if (tools.length === 0) {
                serverToolsList.innerHTML = `<p>${t('server.details.noTools')}</p>`;
            } else {
                let toolsHtml = `<div class="tools-grid">`;

                tools.forEach(tool => {
                    // 説明のエスケープと短縮表示
                    const description = tool.description || t('server.tools.noDescription');
                    const shortDescription = description.length > 100
                        ? description.substring(0, 100) + '...'
                        : description;

                    // パラメータ部分の生成
                    let parametersHtml = '';
                    if (tool.schema && tool.schema.properties) {
                        parametersHtml = `
        <div class="tool-parameters">
            <div class="params-heading">
                <h5>${t('server.tools.parameters')}</h5>
            </div>
            <div class="parameters-details">
                <table class="params-table">
                    <thead>
                        <tr>
                            <th>${t('server.tools.paramName')}</th>
                            <th>${t('server.tools.paramType')}</th>
                            <th>${t('server.tools.paramRequired')}</th>
                            <th>${t('server.tools.paramDescription')}</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${generateParametersTableRows(tool.schema)}
                    </tbody>
                </table>
            </div>
        </div>
    `;
                    } else {
                        parametersHtml = `<div class="tool-parameters">
        <div class="params-heading">
            <h5>${t('server.tools.parameters')}</h5>
        </div>
        <p class="no-params">${t('server.tools.noParams')}</p>
    </div>`;
                    }

                    // ツールカードの生成
                    toolsHtml += `
                        <div class="tool-card">
                            <div class="tool-card-header">
                                <h5 class="tool-name">${tool.name}</h5>
                                <label class="switch small-switch">
                                    <input type="checkbox" class="tool-auto-approve" 
                                           data-server="${server.name}" 
                                           data-tool="${tool.name}" 
                                           ${tool.autoApprove ? 'checked' : ''}>
                                    <span class="slider"></span>
                                </label>
                            </div>
                            
                            <div class="tool-description-container">
                                <div class="tool-description short-desc" id="desc-${tool.name}">
                                    ${shortDescription}
                                </div>
                                <div class="tool-description full-desc hidden" id="full-desc-${tool.name}">
                                    ${description}
                                </div>
                                ${description.length > 100 ?
                            `<button class="btn btn-sm btn-text toggle-desc" data-tool="${tool.name}">
                                        ${t('server.tools.showMore')}
                                    </button>` : ''}
                            </div>
                            
                            ${parametersHtml}
                        </div>
                    `;
                });

                toolsHtml += '</div>';
                serverToolsList.innerHTML = toolsHtml;

                // 説明の展開/折りたたみ処理
                document.querySelectorAll('.toggle-desc').forEach(btn => {
                    btn.addEventListener('click', function () {
                        const toolName = this.getAttribute('data-tool');
                        const shortDesc = document.getElementById(`desc-${toolName}`);
                        const fullDesc = document.getElementById(`full-desc-${toolName}`);

                        shortDesc.classList.toggle('hidden');
                        fullDesc.classList.toggle('hidden');

                        this.textContent = shortDesc.classList.contains('hidden')
                            ? t('server.tools.showLess')
                            : t('server.tools.showMore');
                    });
                });

                // パラメータ表示トグルのイベントリスナー
                document.querySelectorAll('.toggle-params').forEach(btn => {
                    btn.addEventListener('click', function () {
                        const details = this.nextElementSibling;
                        details.classList.toggle('hidden');

                        // ボタンのテキストを切り替え
                        if (details.classList.contains('hidden')) {
                            this.innerHTML = `<span class="material-icons">code</span> ${t('server.tools.showParams')}`;
                        } else {
                            this.innerHTML = `<span class="material-icons">code</span> ${t('server.tools.hideParams')}`;
                        }
                    });
                });

                // 自動承認トグルのイベントリスナー
                document.querySelectorAll('.tool-auto-approve').forEach(toggle => {
                    toggle.addEventListener('change', async (e) => {
                        const serverName = e.target.dataset.server;
                        const toolName = e.target.dataset.tool;
                        const shouldAllow = e.target.checked;

                        try {
                            await toggleToolAutoApprove(serverName, toolName, shouldAllow);
                        } catch (error) {
                            showAlert('danger', t('error.toolToggle', error.message));
                            e.target.checked = !shouldAllow; // 失敗したら元に戻す
                        }
                    });
                });
            }
        } catch (error) {
            toolsLoadingSpinner.style.display = 'none';
            serverToolsList.innerHTML = `<p class="error-message">${t('error.toolsList', error.message)}</p>`;
        }
    }

    // スキーマからパラメーター表示を生成する関数
    function generateParametersTableRows(schema) {
        let rowsHtml = '';

        if (!schema || !schema.properties) {
            return `<tr><td colspan="4">${t('server.tools.noParams')}</td></tr>`;
        }

        const properties = schema.properties;
        const required = schema.required || [];

        // 各プロパティ（パラメータ）の情報を表示
        for (const [name, prop] of Object.entries(properties)) {
            // 型情報の取得
            let type = prop.type || 'any';
            if (prop.enum) {
                type = `enum (${prop.enum.join(', ')})`;
            } else if (type === 'array' && prop.items) {
                type = `array of ${prop.items.type || 'any'}`;
            } else if (type === 'object' && prop.properties) {
                type = 'object';
            }

            // 説明テキストの取得
            const description = prop.description || '';

            // 必須項目か
            const isRequired = required.includes(name);

            rowsHtml += `
                <tr>
                    <td><code>${name}</code></td>
                    <td><code>${type}</code></td>
                    <td>${isRequired ? `<span class="required">${t('server.tools.yes')}</span>` : t('server.tools.no')}</td>
                    <td>${description}</td>
                </tr>
            `;
        }

        return rowsHtml || `<tr><td colspan="4">${t('server.tools.noParams')}</td></tr>`;
    }

    // サーバーのツール一覧を取得
    async function fetchServerTools(serverName) {
        try {
            const response = await fetch(`/admin/servers/${serverName}/tools`);

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || t('error.toolsList', ''));
            }

            const data = await response.json();
            return data.tools || [];
        } catch (error) {
            console.error(t('error.toolsList', ''), error);
            throw error;
        }
    }

    // ツールの自動承認設定を切り替え
    async function toggleToolAutoApprove(serverName, toolName, shouldAllow) {
        const response = await fetch(`/admin/servers/${serverName}/tools/${toolName}/toggleAutoApprove`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ shouldAllow })
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || t('error.toolToggle', ''));
        }

        showAlert(
            'success',
            t('alert.toolToggle.success',
                toolName,
                shouldAllow ? t('alert.toolToggle.enabled') : t('alert.toolToggle.disabled'))
        );
    }

    // 初期データの読み込み
    fetchAndRenderServers();
});