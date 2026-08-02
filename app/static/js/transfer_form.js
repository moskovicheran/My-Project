// Shared autocomplete for the money-transfer forms (admin / agent / club).
// Replaces the from/to <select> dropdowns with type-to-search inputs that
// reveal matching players only after 2 typed characters. Mirrors the
// .sa-autocomplete widget used on the agents page.
//
// Each form must provide, per side (prefix = 'from' / 'to'):
//   #<prefix>Search    text input the user types into
//   #<prefix>Key       hidden input named from_key / to_key ("player_id|nickname")
//   #<prefix>Dropdown  empty .sa-autocomplete container
// and (optionally) #fromBalance / #toBalance / #amountInput for the limit hints.
(function () {
    var selected = { from: null, to: null };  // chosen player's balance (abs)

    function esc(s) {
        return String(s).replace(/[&<>"']/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
        });
    }

    // Rank matches so names that START with the query come first (then any
    // substring match), so a wanted player isn't buried past the cap when a
    // short query hits many names. Caps at 15.
    function matchPlayers(players, q) {
        q = (q || '').trim().toLowerCase();
        var pre = [], sub = [];
        for (var i = 0; i < players.length; i++) {
            var at = players[i].label.toLowerCase().indexOf(q);
            if (at === 0) pre.push(players[i]);
            else if (at > 0) sub.push(players[i]);
        }
        return pre.concat(sub).slice(0, 15);
    }

    function wire(prefix, players) {
        var search = document.getElementById(prefix + 'Search');
        var hidden = document.getElementById(prefix + 'Key');
        var dd = document.getElementById(prefix + 'Dropdown');
        if (!search || !hidden || !dd) return;

        function close() { dd.style.display = 'none'; }

        function render() {
            var q = search.value.trim().toLowerCase();
            // Typing invalidates any prior pick until a match is clicked.
            hidden.value = '';
            selected[prefix] = null;
            updateLimits();
            if (q.length < 2) { close(); return; }
            var matches = matchPlayers(players, q);
            if (!matches.length) { close(); return; }
            dd.innerHTML = matches.map(function (p) {
                return '<div class="sa-item" data-key="' + esc(p.key) +
                       '" data-balance="' + p.balance + '">' + esc(p.label) + '</div>';
            }).join('');
            dd.style.display = 'block';
            Array.prototype.forEach.call(dd.querySelectorAll('.sa-item'), function (item) {
                item.addEventListener('click', function () {
                    hidden.value = item.getAttribute('data-key');
                    search.value = item.textContent;
                    selected[prefix] = parseFloat(item.getAttribute('data-balance'));
                    close();
                    updateLimits();
                });
            });
        }

        search.addEventListener('input', render);
        search.addEventListener('focus', render);
        document.addEventListener('click', function (e) {
            if (!e.target.closest('#' + prefix + 'Search') &&
                !e.target.closest('#' + prefix + 'Dropdown')) close();
        });
    }

    function bal(v) {
        return '<strong style="color:' + (v < 0 ? '#ef233c' : '#2ec4b6') + ';">' + v.toFixed(2) + '</strong>';
    }

    function clubInvolved() {
        // Club wallets (__club__<name>) are free-moving internal buckets, so
        // the smart cap / minus→minus rule doesn't apply when one is picked.
        var f = document.getElementById('fromKey');
        var t = document.getElementById('toKey');
        return (f && f.value.indexOf('__club__') === 0) ||
               (t && t.value.indexOf('__club__') === 0);
    }

    function updateLimits() {
        // Smart cap (never lets anyone go into minus):
        //   payer in plus  → he gives, max = his balance.
        //   payer in minus → he settles a debt, max = min(|debt|, receiver's credit).
        var p = selected.from;   // payer balance (signed)
        var r = selected.to;     // receiver balance (signed)
        var fb = document.getElementById('fromBalance');
        var tb = document.getElementById('toBalance');
        var amt = document.getElementById('amountInput');
        if (fb) fb.innerHTML = (p !== null) ? 'יתרה: ' + bal(p) : '';
        if (tb) tb.innerHTML = (r !== null) ? 'יתרה: ' + bal(r) : '';
        if (amt) {
            if (clubInvolved()) { amt.removeAttribute('max'); }
            else if (p !== null) {
                var max = 0;
                if (p > 0) max = p;
                else if (p < 0 && r !== null && r > 0) max = Math.min(Math.abs(p), r);
                amt.max = max > 0 ? max : '';
            }
        }
    }

    // Block submit until both sides were actually picked from the list, and
    // stop minus→minus (a player in debt can only pay someone in plus).
    window.validateTransfer = function () {
        var from = document.getElementById('fromKey');
        var to = document.getElementById('toKey');
        if (!from || !from.value || !to || !to.value) {
            alert('יש לבחור משלם ומקבל מהרשימה.');
            return false;
        }
        if (!clubInvolved() &&
            selected.from !== null && selected.from < 0 &&
            selected.to !== null && selected.to <= 0) {
            alert('לא ניתן להעביר ממינוס למינוס — שחקן בחוב יכול להעביר רק לשחקן בפלוס.');
            return false;
        }
        return true;
    };

    window.initTransferForm = function (fromPlayers, toPlayers) {
        wire('from', fromPlayers);
        wire('to', toPlayers);
    };

    // Generic single-field player autocomplete (house return / distribute).
    // prefix drives the element ids: <prefix>Search/Key/Dropdown/Balance/Amount.
    // capFromBalance=true caps the amount input at the picked player's balance
    // (return-to-house); false leaves the amount cap to the form (distribute,
    // capped server-side and via the input's own max = house pot).
    function single(prefix, players, capFromBalance) {
        var search = document.getElementById(prefix + 'Search');
        var hidden = document.getElementById(prefix + 'Key');
        var dd = document.getElementById(prefix + 'Dropdown');
        var hint = document.getElementById(prefix + 'Balance');
        var amt = document.getElementById(prefix + 'Amount');
        if (!search || !hidden || !dd) return;
        function close() { dd.style.display = 'none'; }
        function render() {
            var q = search.value.trim().toLowerCase();
            hidden.value = '';
            if (hint) hint.innerHTML = '';
            if (capFromBalance && amt) amt.removeAttribute('max');
            if (q.length < 2) { close(); return; }
            var matches = matchPlayers(players, q);
            if (!matches.length) { close(); return; }
            dd.innerHTML = matches.map(function (p) {
                return '<div class="sa-item" data-key="' + esc(p.key) +
                       '" data-balance="' + p.balance + '">' + esc(p.label) + '</div>';
            }).join('');
            dd.style.display = 'block';
            Array.prototype.forEach.call(dd.querySelectorAll('.sa-item'), function (item) {
                item.addEventListener('click', function () {
                    hidden.value = item.getAttribute('data-key');
                    search.value = item.textContent;
                    var b = parseFloat(item.getAttribute('data-balance'));
                    if (hint) hint.innerHTML = 'יתרה: ' + bal(b);
                    if (capFromBalance && amt && b > 0) amt.max = b;
                    close();
                });
            });
        }
        search.addEventListener('input', render);
        search.addEventListener('focus', render);
        document.addEventListener('click', function (e) {
            if (!e.target.closest('#' + prefix + 'Search') &&
                !e.target.closest('#' + prefix + 'Dropdown')) close();
        });
    }

    // Return-to-house works for any player (incl. minus) and isn't capped by
    // the player's balance — it's a correction — so capFromBalance = false.
    window.initReturnHouse = function (players) { single('rh', players, false); };
    window.initDistributeHouse = function (players) { single('dh', players, false); };
    // Player-cross loader (admin balance UI): pick a player, then "load" posts
    // cross_key = "player_id|nickname" to render his per-club breakdown.
    window.initPlayerCross = function (players) { single('cx', players, false); };

    function pickedOrAlert(id) {
        var k = document.getElementById(id);
        if (!k || !k.value) { alert('יש לבחור שחקן מהרשימה.'); return false; }
        return true;
    }
    window.validateReturnHouse = function () { return pickedOrAlert('rhKey'); };
    window.validateDistributeHouse = function () { return pickedOrAlert('dhKey'); };
})();
