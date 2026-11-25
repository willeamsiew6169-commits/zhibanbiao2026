(() => {
    const root = document.documentElement;
    const dock = document.querySelector('.vol-action-dock');
    const viewport = window.visualViewport;
    const mobile = matchMedia('(max-width: 700px), (pointer: coarse)');
    const navigation = document.querySelector('[data-vol-bottom-nav]');
    const updateLayout = () => {
        const bottom = viewport ? Math.max(0, root.clientHeight - viewport.height - viewport.offsetTop) : 0;
        root.style.setProperty('--vol-view-bottom', bottom + 'px');
        root.style.setProperty('--vol-view-top', (viewport ? viewport.offsetTop : 0) + 'px');
        root.style.setProperty('--vol-view-height', viewport ? viewport.height + 'px' : '100dvh');
        root.style.setProperty('--vol-view-left', (viewport ? viewport.offsetLeft : 0) + 'px');
        root.style.setProperty('--vol-view-width', viewport ? viewport.width + 'px' : '100%');
        root.style.setProperty('--vol-nav-height', (navigation ? navigation.offsetHeight : 0) + 'px');
        root.style.setProperty('--vol-dock-height', (dock && mobile.matches ? dock.offsetHeight : 0) + 'px');
    };
    root.classList.add('vol-action-ready');
    if (document.querySelector('meta[name="viewport"][content*="viewport-fit=cover"]')) {
        root.classList.add('vol-safe-area');
    }
    updateLayout();
    window.addEventListener('resize', updateLayout);
    if (viewport) {
        viewport.addEventListener('resize', () => {
            updateLayout();
            const field = document.activeElement;
            if (dock && dock.offsetHeight && field && field.matches('input, textarea, select')
                && field.getBoundingClientRect().bottom > dock.getBoundingClientRect().top) {
                field.scrollIntoView({block: 'center'});
            }
        });
        viewport.addEventListener('scroll', updateLayout);
    }
    if (window.ResizeObserver) {
        const observer = new ResizeObserver(updateLayout);
        if (dock) observer.observe(dock);
        if (navigation) observer.observe(navigation);
    }

    const zh = !root.lang.startsWith('en');
    const setStatus = (node, message) => {
        if (node) node.textContent = message;
        updateLayout();
    };
    const forms = document.querySelectorAll('form[data-vol-submit]');
    forms.forEach(form => {
        const status = form.dataset.volStatus
            ? document.querySelector(form.dataset.volStatus)
            : form.querySelector('.vol-action-status');
        form.addEventListener('submit', async event => {
            if (event.defaultPrevented) return;
            if (form.dataset.volBusy) {
                event.preventDefault();
                return;
            }
            const submitter = event.submitter || document.activeElement;
            if (form.dataset.volOnly && (!submitter || submitter.value !== form.dataset.volOnly)) return;
            event.preventDefault();
            const payload = new FormData(form);
            if (submitter && submitter.form === form && submitter.name) payload.append(submitter.name, submitter.value);
            const linked = form.id ? document.querySelectorAll('[data-vol-lock="' + form.id + '"]') : [];
            const controls = [...form.querySelectorAll('button'), ...linked]
                .filter(button => !button.disabled);
            const unlock = () => {
                delete form.dataset.volBusy;
                controls.forEach(button => { button.disabled = false; });
            };
            form.dataset.volBusy = 'true';
            form.setAttribute('aria-busy', 'true');
            controls.forEach(button => { button.disabled = true; });
            setStatus(status, zh ? '正在处理，请稍候' : 'Processing, please wait');
            const controller = window.AbortController ? new AbortController() : null;
            const timer = controller && setTimeout(() => controller.abort(), 90000);
            try {
                const response = await fetch(form.action, {
                    method: 'POST',
                    headers: {'X-Volunteer-Submit': '1'},
                    body: new URLSearchParams(payload),
                    signal: controller ? controller.signal : undefined
                });
                if (response.status >= 500 || new URL(response.url).origin !== location.origin) throw new Error();
                const type = response.headers.get('Content-Type') || '';
                if (response.redirected || (response.ok && type.includes('application/json'))) {
                    const next = new URL(response.redirected ? response.url : (await response.json()).next, location.href);
                    if (next.origin !== location.origin) throw new Error();
                    location.assign(next.href);
                    return;
                }
                const resultPage = type.includes('text/html')
                    ? new DOMParser().parseFromString(await response.text(), 'text/html')
                    : null;
                const result = resultPage && resultPage.querySelector('[data-vol-result]');
                if (!response.ok && !result) {
                    setStatus(status, zh ? '操作失败，尚未完成。请检查资料或重新进入页面。' : 'Operation failed. Check your details or reopen the page.');
                    unlock();
                    return;
                }
                if (!result) throw new Error();
                if (result.dataset.volResult === 'retry') {
                    setStatus(status, result.textContent.trim());
                    unlock();
                    return;
                }
                // These receipts are static. The signed-in homepage navigates normally.
                const activeField = document.activeElement;
                if (result.hasAttribute('data-vol-blur') && activeField && activeField.form === form && activeField.matches('input')) {
                    activeField.blur();
                }
                resultPage.querySelectorAll('script').forEach(script => script.remove());
                document.replaceChild(document.adoptNode(resultPage.documentElement), root);
                result.scrollIntoView({block: 'nearest'});
            } catch (error) {
                const next = form.dataset.volRecords ? (zh ? '请先查询记录。' : 'Check your records first.')
                    : (zh ? '请先向现场负责人查询记录。' : 'Ask the onsite coordinator to check your record first.');
                setStatus(status, (zh ? '网络异常，暂时无法确认结果。' : 'Connection problem. The result is unknown. ') + next);
                if (status && form.dataset.volRecords) {
                    const records = document.createElement('a');
                    records.href = form.dataset.volRecords;
                    records.textContent = zh ? '查询记录' : 'Check records';
                    status.appendChild(records);
                    updateLayout();
                }
            } finally {
                clearTimeout(timer);
                form.removeAttribute('aria-busy');
            }
        });
    });
})();
