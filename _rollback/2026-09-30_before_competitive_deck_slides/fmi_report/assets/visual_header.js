/* Make the per-visual Focus icon usable:
   ⤢ (title "Focus") -> enlarge the panel into a centered focus overlay.
   One document-level listener handles every panel on every page, so it
   works for the dynamically-rendered visuals too. */
(function () {
    function fireResize() { window.dispatchEvent(new Event("resize")); }

    function closeFocus() {
        var f = document.querySelector(".vp-focused");
        if (f) {
            f.classList.remove("vp-focused");
            var c = f.querySelector(".vp-close");
            if (c) c.remove();
        }
        var bd = document.getElementById("viz-backdrop");
        if (bd) bd.remove();
        fireResize();                       // let Plotly re-fit
    }

    function focusPanel(panel) {
        if (panel.classList.contains("vp-focused")) { closeFocus(); return; }
        closeFocus();
        var bd = document.createElement("div");
        bd.id = "viz-backdrop";
        bd.addEventListener("click", closeFocus);
        document.body.appendChild(bd);
        panel.classList.add("vp-focused");
        var close = document.createElement("div");
        close.className = "vp-close";
        close.textContent = "✕";
        close.title = "Close";
        close.addEventListener("click", function (ev) {
            ev.stopPropagation(); closeFocus();
        });
        panel.appendChild(close);
        setTimeout(fireResize, 80);
    }

    document.addEventListener("click", function (e) {
        var icon = e.target.closest(".vh-icon");
        if (icon && icon.getAttribute("title") === "Focus") {
            e.stopPropagation();
            var panel = icon.closest(".viz-panel");
            if (panel) focusPanel(panel);
        }
    });

    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape") { closeFocus(); }
    });
})();
