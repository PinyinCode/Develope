# -*- coding: utf-8 -*-
"""
Template cho tính năng GIA HẠN + TRIAL tự động.
- Đăng ký lần đầu → tặng 7 ngày (ATOMIC transaction, chống race condition)
- User chọn gói → tạo renewal_request → hiện QR ngân hàng
- User bấm "Đã thanh toán" → status user_paid
- Admin xác nhận thủ công → status confirmed
- Listener Firestore tự cập nhật UI user
KHÔNG CẦN SỬA khi đổi cấu trúc Excel hay giao diện học.
"""


def build_renewal_css():
    return r"""
/* ============ RENEWAL MODAL ============ */
.renewal-modal{
    position:fixed;inset:0;background:rgba(15,23,42,.85);
    backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);
    z-index:3500;display:none;align-items:center;justify-content:center;
    padding:1rem;animation:fadeIn .2s;overflow-y:auto;
}
.renewal-modal.show{display:flex}
.renewal-box{
    background:var(--surface);border-radius:20px;
    width:100%;max-width:560px;max-height:calc(100vh - 2rem);
    box-shadow:0 20px 60px rgba(0,0,0,.4);
    display:flex;flex-direction:column;overflow:hidden;
    animation:slideUp .3s cubic-bezier(.34,1.56,.64,1);
}
.renewal-header{
    padding:1.25rem 1.5rem;border-bottom:1px solid var(--border);
    display:flex;align-items:center;justify-content:space-between;gap:1rem;
    background:linear-gradient(135deg,#2563eb,#7c3aed);color:#fff;
}
.renewal-header h2{
    font-size:1.15rem;font-weight:800;display:flex;align-items:center;gap:.5rem;
    color:#fff;margin:0;
}
.renewal-header .subtitle{font-size:.78rem;opacity:.9;margin-top:.15rem;}
.renewal-close{
    width:34px;height:34px;border-radius:50%;border:none;
    background:rgba(255,255,255,.2);color:#fff;cursor:pointer;
    display:flex;align-items:center;justify-content:center;
    font-size:1rem;transition:.15s;flex-shrink:0;
}
.renewal-close:hover{background:rgba(255,255,255,.35)}
.renewal-body{padding:1.5rem;overflow-y:auto;flex:1;}

.renewal-current{
    padding:.85rem 1rem;border-radius:12px;
    background:linear-gradient(135deg, rgba(37,99,235,.08), rgba(124,58,237,.08));
    border:1px solid rgba(37,99,235,.25);
    margin-bottom:1.25rem;
    display:flex;align-items:center;gap:.75rem;
}
.renewal-current .rc-icon{
    width:42px;height:42px;border-radius:12px;
    background:var(--primary);color:#fff;
    display:flex;align-items:center;justify-content:center;
    font-size:1.15rem;flex-shrink:0;
}
.renewal-current.warn .rc-icon{background:var(--amber);}
.renewal-current.expired .rc-icon{background:var(--danger);}
.renewal-current .rc-info{flex:1;min-width:0;}
.renewal-current .rc-title{
    font-weight:800;font-size:.9rem;color:var(--text);margin-bottom:.15rem;
}
.renewal-current .rc-desc{font-size:.75rem;color:var(--text-2);line-height:1.4;}
.renewal-current .rc-desc b{color:var(--primary);}

.renewal-section-title{
    font-size:.72rem;font-weight:800;color:var(--text-3);
    text-transform:uppercase;letter-spacing:.5px;margin-bottom:.6rem;
    display:flex;align-items:center;gap:.35rem;
}
.package-grid{
    display:grid;grid-template-columns:1fr;gap:.6rem;margin-bottom:1.25rem;
}
@media(min-width:480px){.package-grid{grid-template-columns:1fr 1fr 1fr;}}
.package-card{
    padding:.9rem .75rem;border-radius:12px;
    border:2px solid var(--border);background:var(--surface);
    cursor:pointer;transition:.2s;text-align:center;
    position:relative;user-select:none;
}
.package-card:hover{border-color:var(--primary);transform:translateY(-2px);}
.package-card.selected{
    border-color:var(--primary);
    background:linear-gradient(135deg, rgba(37,99,235,.08), rgba(124,58,237,.08));
    box-shadow:0 6px 20px rgba(37,99,235,.25);
}
.package-card .pkg-label{
    font-weight:800;font-size:.95rem;color:var(--text);margin-bottom:.3rem;
}
.package-card .pkg-price{
    font-weight:900;font-size:1.25rem;color:var(--primary);
    line-height:1;margin-bottom:.25rem;
}
.package-card .pkg-unit{font-size:.68rem;color:var(--text-3);font-weight:600;}
.package-card .pkg-save{
    position:absolute;top:-8px;right:-4px;
    background:linear-gradient(135deg,#16a34a,#22c55e);
    color:#fff;font-size:.6rem;font-weight:800;
    padding:.15rem .45rem;border-radius:50px;
    text-transform:uppercase;letter-spacing:.3px;
    box-shadow:0 2px 6px rgba(22,163,74,.4);
}
.package-card .pkg-popular{
    position:absolute;top:-8px;left:50%;transform:translateX(-50%);
    background:linear-gradient(135deg,#f59e0b,#d97706);
    color:#fff;font-size:.6rem;font-weight:800;
    padding:.15rem .5rem;border-radius:50px;
    text-transform:uppercase;letter-spacing:.3px;white-space:nowrap;
    box-shadow:0 2px 6px rgba(245,158,11,.4);
}

.payment-box{display:none;flex-direction:column;gap:1rem;animation:fadeIn .3s;}
.payment-box.show{display:flex;}
.qr-wrap{
    display:flex;flex-direction:column;align-items:center;gap:.75rem;
    padding:1rem;background:linear-gradient(135deg,#f0f4f8,#e2e8f0);
    border-radius:14px;border:1.5px dashed var(--border);
}
[data-theme="dark"] .qr-wrap{
    background:linear-gradient(135deg,#1e293b,#0f172a);border-color:#475569;
}
.qr-img{
    width:220px;height:220px;background:#fff;padding:.5rem;border-radius:10px;
    box-shadow:0 4px 16px rgba(0,0,0,.15);
}
.qr-wrap .qr-hint{
    font-size:.75rem;color:var(--text-2);text-align:center;
    line-height:1.5;max-width:300px;
}
.qr-wrap .qr-hint b{color:var(--primary);}

.bank-info{
    padding:.9rem 1rem;border-radius:12px;
    background:var(--surface-2);border:1px solid var(--border);
    display:flex;flex-direction:column;gap:.55rem;
}
.bank-row{
    display:flex;justify-content:space-between;align-items:center;gap:.75rem;
    font-size:.82rem;
}
.bank-row .br-label{color:var(--text-3);font-weight:600;flex-shrink:0;}
.bank-row .br-value{
    color:var(--text);font-weight:700;
    word-break:break-all;text-align:right;
}
.bank-row .br-value.code{
    font-family:ui-monospace,SFMono-Regular,Menlo,monospace;
    color:var(--danger);font-size:.95rem;
    background:rgba(220,38,38,.08);padding:.15rem .5rem;border-radius:6px;
    letter-spacing:1px;
}
.copy-btn{
    width:28px;height:28px;border-radius:6px;border:1px solid var(--border);
    background:var(--surface);color:var(--text-2);cursor:pointer;
    display:inline-flex;align-items:center;justify-content:center;
    font-size:.75rem;transition:.15s;margin-left:.4rem;
}
.copy-btn:hover{background:var(--primary-light);color:var(--primary-dark);border-color:var(--primary);}
.copy-btn.copied{background:var(--success);color:#fff;border-color:var(--success);}

.payment-steps{
    display:flex;flex-direction:column;gap:.5rem;
    padding:.85rem 1rem;border-radius:12px;
    background:linear-gradient(135deg, rgba(245,158,11,.1), rgba(245,158,11,.05));
    border:1px solid rgba(245,158,11,.3);
}
.payment-steps .step{
    display:flex;gap:.6rem;font-size:.8rem;color:var(--text-2);
    line-height:1.5;align-items:flex-start;
}
.payment-steps .step .num{
    width:20px;height:20px;border-radius:50%;
    background:var(--amber);color:#fff;
    display:flex;align-items:center;justify-content:center;
    font-size:.68rem;font-weight:800;flex-shrink:0;margin-top:.05rem;
}
.payment-steps .step b{color:var(--text);}

.renewal-actions{
    display:flex;gap:.6rem;justify-content:flex-end;
    padding-top:1rem;border-top:1px solid var(--border);
}
.renewal-btn{
    padding:.7rem 1.2rem;border-radius:10px;border:1.5px solid var(--border);
    background:var(--surface);color:var(--text);
    font-size:.85rem;font-weight:700;cursor:pointer;
    transition:.15s;font-family:inherit;
    display:inline-flex;align-items:center;gap:.4rem;
}
.renewal-btn:hover{background:var(--surface-2);}
.renewal-btn.primary{
    background:var(--primary);color:#fff;border-color:var(--primary);
    box-shadow:0 4px 12px rgba(37,99,235,.3);
}
.renewal-btn.primary:hover{background:var(--primary-dark);}
.renewal-btn.success{
    background:var(--success);color:#fff;border-color:var(--success);
    box-shadow:0 4px 12px rgba(22,163,74,.3);
}
.renewal-btn.success:hover{background:#15803d;}
.renewal-btn:disabled{opacity:.5;cursor:not-allowed;}

.renewal-success{
    text-align:center;padding:2rem 1rem;
    display:flex;flex-direction:column;align-items:center;gap:1rem;
}
.renewal-success .icon{
    width:70px;height:70px;border-radius:50%;
    background:linear-gradient(135deg,#16a34a,#22c55e);
    color:#fff;display:flex;align-items:center;justify-content:center;
    font-size:2rem;box-shadow:0 8px 24px rgba(22,163,74,.4);
    animation:successPop .5s cubic-bezier(.34,1.56,.64,1);
}
@keyframes successPop{
    0%{transform:scale(0);}
    70%{transform:scale(1.15);}
    100%{transform:scale(1);}
}
.renewal-success h3{font-size:1.15rem;font-weight:800;color:var(--text);margin:0;}
.renewal-success p{
    font-size:.85rem;color:var(--text-2);line-height:1.6;
    max-width:340px;margin:0;
}
.renewal-success .info-box{
    padding:.65rem 1rem;border-radius:10px;
    background:var(--surface-2);border:1px solid var(--border);
    font-size:.78rem;color:var(--text-2);
    display:flex;align-items:center;gap:.5rem;text-align:left;
}

.trial-banner{
    background:linear-gradient(135deg,#dbeafe,#bfdbfe);
    border:1.5px solid #2563eb;border-radius:var(--radius);
    padding:.9rem 1.1rem;margin-bottom:1rem;
    display:flex;align-items:center;gap:.75rem;flex-wrap:wrap;
}
[data-theme="dark"] .trial-banner{
    background:linear-gradient(135deg,rgba(37,99,235,.2),rgba(37,99,235,.15));
    border-color:#3b82f6;
}
.trial-banner-icon{
    width:36px;height:36px;border-radius:50%;
    background:#2563eb;color:#fff;
    display:flex;align-items:center;justify-content:center;font-size:1rem;flex-shrink:0;
}
.trial-banner-text{flex:1;min-width:200px}
.trial-banner-text .title{
    font-weight:700;font-size:.9rem;color:#1e40af;margin-bottom:.15rem;
}
[data-theme="dark"] .trial-banner-text .title{color:#93c5fd;}
.trial-banner-text .desc{font-size:.78rem;color:#1e40af;line-height:1.5;}
[data-theme="dark"] .trial-banner-text .desc{color:#bfdbfe;}
.trial-banner-text b{color:#dc2626;}
.trial-banner-btn{
    padding:.5rem .9rem;border-radius:50px;border:none;
    background:#2563eb;color:#fff;text-decoration:none;
    font-size:.8rem;font-weight:700;cursor:pointer;transition:.15s;
    display:inline-flex;align-items:center;gap:.35rem;white-space:nowrap;
    font-family:inherit;
}
.trial-banner-btn:hover{background:#1d4ed8;color:#fff;transform:translateY(-1px);}

.dropdown-renew{
    display:flex;align-items:center;gap:.5rem;width:100%;
    padding:.65rem .75rem;border-radius:var(--radius-sm);
    background:linear-gradient(135deg,#f59e0b,#d97706);
    color:#fff !important;font-size:.85rem;font-weight:700;
    text-decoration:none;transition:.15s;margin-top:.25rem;
    border:none;cursor:pointer;font-family:inherit;text-align:left;
    box-shadow:0 2px 8px rgba(245,158,11,.3);
}
.dropdown-renew:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(245,158,11,.4);color:#fff !important;}
.dropdown-renew i{font-size:1rem;}

.renewal-admin-row{
    display:flex;flex-direction:column;gap:.5rem;
    padding:.85rem;border-radius:10px;
    background:var(--surface-2);border:1px solid var(--border);
    margin-bottom:.5rem;
}
.renewal-admin-row .rar-head{
    display:flex;justify-content:space-between;gap:.5rem;flex-wrap:wrap;
    align-items:flex-start;
}
.renewal-admin-row .rar-email{
    font-weight:800;font-size:.88rem;color:var(--text);
    word-break:break-all;
}
.renewal-admin-row .rar-sub{
    font-size:.72rem;color:var(--text-3);margin-top:.15rem;
    display:flex;align-items:center;gap:.3rem;flex-wrap:wrap;
}
.renewal-admin-row .rar-pkg{text-align:right;}
.renewal-admin-row .rar-amount{
    font-weight:900;font-size:1.05rem;color:var(--primary);
    line-height:1;
}
.renewal-admin-row .rar-pkg-label{
    font-size:.7rem;color:var(--text-3);margin-top:.15rem;
}
.renewal-admin-row .rar-foot{
    display:flex;justify-content:space-between;align-items:center;gap:.5rem;
    padding-top:.5rem;border-top:1px solid var(--border);
    flex-wrap:wrap;
}
.renewal-admin-row .rar-code{
    font-family:ui-monospace,SFMono-Regular,Menlo,monospace;
    font-size:.78rem;color:var(--danger);
    background:rgba(220,38,38,.08);
    padding:.15rem .5rem;border-radius:6px;
    letter-spacing:.5px;font-weight:800;
}
.renewal-admin-row .rar-status{
    display:inline-flex;align-items:center;gap:.25rem;
    padding:.2rem .5rem;border-radius:50px;
    font-size:.65rem;font-weight:800;
    text-transform:uppercase;letter-spacing:.3px;
}
.renewal-admin-row .rar-status.pending{background:rgba(37,99,235,.15);color:#2563eb;}
.renewal-admin-row .rar-status.user_paid{background:rgba(245,158,11,.18);color:#d97706;}
.renewal-admin-row .rar-actions{
    display:flex;gap:.35rem;align-items:center;flex-wrap:wrap;
}
"""


def build_renewal_html():
    return r"""
<div class="renewal-modal" id="renewalModal">
    <div class="renewal-box">
        <div class="renewal-header">
            <div>
                <h2><i class="fas fa-crown"></i> Gia hạn tài khoản</h2>
                <div class="subtitle">Chọn gói phù hợp và thanh toán</div>
            </div>
            <button class="renewal-close" id="renewalClose"><i class="fas fa-times"></i></button>
        </div>
        <div class="renewal-body" id="renewalBody">
            <!-- JS render -->
        </div>
    </div>
</div>
"""


def build_renewal_js():
    return r"""
/* ═══════════════════════════════════════════════════════════════
   GIA HẠN TÀI KHOẢN - TRIAL TỰ ĐỘNG + XÁC NHẬN THỦ CÔNG
   ═══════════════════════════════════════════════════════════════ */

var TRIAL_DAYS = __TRIAL_DAYS__;
var BANK_CONFIG = __BANK_CONFIG__;
var PACKAGES = __PACKAGES__;
var RENEWAL_SUPPORT_ZALO = "__RENEWAL_SUPPORT_ZALO__";

var renewalSelectedPkg = null;
var renewalCurrentReq = null;
var renewalListener = null;

/* ─── ĐĂNG KÝ TRIAL: TỰ ĐỘNG TẶNG 7 NGÀY (ATOMIC) ─────────────── */
async function grantTrialIfNew(user, userData) {
    if (!user || !user.email) return false;

    var email = user.email.toLowerCase();
    var userRef = db.collection('allowed_users').doc(email);

    try {
        /* ✅ Dùng transaction để tránh race condition (2 tab login cùng lúc) */
        var result = await db.runTransaction(async function(transaction) {
            var doc = await transaction.get(userRef);

            if (doc.exists) {
                var data = doc.data() || {};
                /* Doc tồn tại và đã có dấu hiệu đăng ký → KHÔNG tặng trial */
                if (data.registeredAt || data.expiresAt || data.role === 'admin') {
                    return { granted: false, reason: 'already_exists' };
                }
                /* Doc tồn tại nhưng hoàn toàn rỗng → vẫn coi như user mới */
            }

            var trialMs = TRIAL_DAYS * 24 * 60 * 60 * 1000;
            var expiresAt = new Date(Date.now() + trialMs);
            expiresAt.setHours(23, 59, 59, 0);

            var setData = {
                name: (userData && userData.name)
                    ? userData.name
                    : (user.displayName || email.split('@')[0]),
                role: 'user',
                expiresAt: firebase.firestore.Timestamp.fromDate(expiresAt),
                registeredAt: firebase.firestore.FieldValue.serverTimestamp(),
                isTrial: true,
                trialDays: TRIAL_DAYS,
                trialStartedAt: firebase.firestore.FieldValue.serverTimestamp()
            };

            transaction.set(userRef, setData, { merge: false });

            return { granted: true, expiresAt: expiresAt };
        });

        if (result.granted) {
            console.log('✅ Trial granted:', TRIAL_DAYS, 'days for', email);
        } else {
            console.log('ℹ️ Trial skipped (already exists):', email);
        }
        return result.granted;
    } catch (e) {
        console.error('❌ Grant trial error:', e);
        return false;
    }
}

/* ─── MỞ MODAL GIA HẠN ────────────────────────────────────────── */
window.openRenewalModal = function() {
    if (!currentUser) {
        showLoginModal();
        return;
    }
    if (currentUser.role === 'admin') {
        alert('Admin có hạn vĩnh viễn, không cần gia hạn!');
        return;
    }
    renewalSelectedPkg = null;
    renewalCurrentReq = null;
    renderRenewalStep1();
    $('renewalModal').classList.add('show');
    if ($('userDropdown')) $('userDropdown').classList.remove('show');
};

window.closeRenewalModal = function() {
    $('renewalModal').classList.remove('show');
    if (renewalListener) {
        try { renewalListener(); } catch(e) {}
        renewalListener = null;
    }
};

/* ─── BƯỚC 1: CHỌN GÓI ────────────────────────────────────────── */
function renderRenewalStep1() {
    var daysLeft = getDaysRemaining(currentUser);
    var isExpired = daysLeft !== null && daysLeft <= 0;
    var isWarn = daysLeft !== null && daysLeft > 0 && daysLeft <= 7;

    var currentCls = isExpired ? 'expired' : (isWarn ? 'warn' : '');
    var currentIcon = isExpired ? 'fa-exclamation-triangle'
                     : (isWarn ? 'fa-hourglass-half' : 'fa-calendar-check');
    var currentTitle = isExpired ? 'Tài khoản đã hết hạn'
                     : (isWarn ? 'Sắp hết hạn' : 'Tài khoản đang hoạt động');
    var currentDesc = '';
    if (daysLeft === null) {
        currentDesc = 'Vĩnh viễn, không cần gia hạn';
    } else if (isExpired) {
        currentDesc = 'Đã hết hạn <b>' + Math.abs(daysLeft) +
                      ' ngày</b> trước. Gia hạn ngay để tiếp tục học!';
    } else {
        currentDesc = 'Còn <b>' + daysLeft +
                      ' ngày</b> sử dụng. Gia hạn để không bị gián đoạn!';
    }

    var packagesHtml = '';
    PACKAGES.forEach(function(p) {
        var saveHtml = p.save ? '<div class="pkg-save">' + escapeHtml(p.save) + '</div>' : '';
        var popularHtml = p.popular ? '<div class="pkg-popular">⭐ Phổ biến</div>' : '';
        packagesHtml +=
            '<div class="package-card" data-pkg="' + p.id +
                '" onclick="selectPackage(\'' + p.id + '\')">' +
                popularHtml + saveHtml +
                '<div class="pkg-label">' + escapeHtml(p.label) + '</div>' +
                '<div class="pkg-price">' + formatMoney(p.amount) + '</div>' +
                '<div class="pkg-unit">VNĐ</div>' +
            '</div>';
    });

    $('renewalBody').innerHTML =
        '<div class="renewal-current ' + currentCls + '">' +
            '<div class="rc-icon"><i class="fas ' + currentIcon + '"></i></div>' +
            '<div class="rc-info">' +
                '<div class="rc-title">' + currentTitle + '</div>' +
                '<div class="rc-desc">' + currentDesc + '</div>' +
            '</div>' +
        '</div>' +

        '<div class="renewal-section-title"><i class="fas fa-box"></i> Chọn gói gia hạn</div>' +
        '<div class="package-grid">' + packagesHtml + '</div>' +

        '<div class="renewal-actions">' +
            '<button class="renewal-btn" onclick="closeRenewalModal()">Hủy</button>' +
            '<button class="renewal-btn primary" id="renewalNextBtn" disabled onclick="goToPayment()">' +
                '<i class="fas fa-arrow-right"></i> Tiếp tục' +
            '</button>' +
        '</div>';
}

window.selectPackage = function(pkgId) {
    renewalSelectedPkg = PACKAGES.find(function(p) { return p.id === pkgId; });
    document.querySelectorAll('.package-card').forEach(function(c) {
        c.classList.toggle('selected', c.dataset.pkg === pkgId);
    });
    var btn = $('renewalNextBtn');
    if (btn) btn.disabled = false;
};

/* ─── BƯỚC 2: THANH TOÁN (QR + STK) ────────────────────────────── */
async function goToPayment() {
    if (!renewalSelectedPkg) return;

    var transferCode = generateTransferCode(currentUser.email);
    var amount = renewalSelectedPkg.amount;

    var btn = $('renewalNextBtn');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Đang tạo...';
    }

    try {
        var reqRef = await db.collection('renewal_requests').add({
            email: currentUser.email,
            name: currentUser.name,
            package: renewalSelectedPkg.id,
            packageLabel: renewalSelectedPkg.label,
            amount: amount,
            days: renewalSelectedPkg.days,
            transferCode: transferCode,
            status: 'pending',
            method: 'manual',
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
        });
        renewalCurrentReq = { id: reqRef.id, code: transferCode };
    } catch(e) {
        alert('❌ Lỗi tạo yêu cầu: ' + e.message);
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-arrow-right"></i> Tiếp tục';
        }
        return;
    }

    renderPaymentScreen();
    listenRenewalRequest(reqRef.id);
}

function renderPaymentScreen() {
    var pkg = renewalSelectedPkg;
    var code = renewalCurrentReq.code;
    var amount = pkg.amount;

    var qrUrl = 'https://img.vietqr.io/image/' + BANK_CONFIG.bank_id +
                '-' + BANK_CONFIG.account_no + '-compact2.png' +
                '?amount=' + amount +
                '&addInfo=' + encodeURIComponent(code) +
                '&accountName=' + encodeURIComponent(BANK_CONFIG.account_name);

    $('renewalBody').innerHTML =
        '<div class="renewal-current">' +
            '<div class="rc-icon"><i class="fas fa-shopping-cart"></i></div>' +
            '<div class="rc-info">' +
                '<div class="rc-title">Gói ' + escapeHtml(pkg.label) + '</div>' +
                '<div class="rc-desc">Số tiền: <b>' + formatMoney(amount) +
                    ' VNĐ</b> · ' + pkg.days + ' ngày</div>' +
            '</div>' +
        '</div>' +

        '<div class="renewal-section-title"><i class="fas fa-qrcode"></i> Quét mã để thanh toán</div>' +
        '<div class="qr-wrap">' +
            '<img class="qr-img" src="' + qrUrl + '" alt="QR thanh toán" ' +
                'onerror="this.style.display=\'none\'">' +
            '<div class="qr-hint">' +
                'Mở app ngân hàng, quét mã QR để chuyển khoản.<br>' +
                'Hoặc chuyển thủ công theo thông tin bên dưới.' +
            '</div>' +
        '</div>' +

        '<div class="bank-info">' +
            '<div class="bank-row">' +
                '<span class="br-label">Ngân hàng</span>' +
                '<span class="br-value">' + escapeHtml(BANK_CONFIG.bank_name) + '</span>' +
            '</div>' +
            '<div class="bank-row">' +
                '<span class="br-label">Số TK</span>' +
                '<span class="br-value">' + escapeHtml(BANK_CONFIG.account_no) +
                    '<button class="copy-btn" onclick="copyText(\'' +
                        escapeJs(BANK_CONFIG.account_no) +
                        '\', this)"><i class="fas fa-copy"></i></button>' +
                '</span>' +
            '</div>' +
            '<div class="bank-row">' +
                '<span class="br-label">Chủ TK</span>' +
                '<span class="br-value">' + escapeHtml(BANK_CONFIG.account_name) + '</span>' +
            '</div>' +
            '<div class="bank-row">' +
                '<span class="br-label">Số tiền</span>' +
                '<span class="br-value code">' + formatMoney(amount) + 'đ</span>' +
            '</div>' +
            '<div class="bank-row">' +
                '<span class="br-label">Nội dung</span>' +
                '<span class="br-value code">' + escapeHtml(code) +
                    '<button class="copy-btn" onclick="copyText(\'' +
                        escapeJs(code) +
                        '\', this)"><i class="fas fa-copy"></i></button>' +
                '</span>' +
            '</div>' +
        '</div>' +

        '<div class="payment-steps">' +
            '<div class="step"><span class="num">1</span>' +
                '<div>Mở app ngân hàng, chuyển <b>' + formatMoney(amount) +
                ' VNĐ</b> đến STK trên</div></div>' +
            '<div class="step"><span class="num">2</span>' +
                '<div>Ghi đúng nội dung: <b>' + escapeHtml(code) + '</b></div></div>' +
            '<div class="step"><span class="num">3</span>' +
                '<div>Nhấn nút <b>"Tôi đã thanh toán"</b> bên dưới</div></div>' +
            '<div class="step"><span class="num">4</span>' +
                '<div>Hệ thống xác nhận trong <b>1-5 phút</b>, ' +
                'tài khoản tự động gia hạn</div></div>' +
        '</div>' +

        '<div class="renewal-actions">' +
            '<button class="renewal-btn" onclick="cancelRenewal()">' +
                '<i class="fas fa-times"></i> Hủy' +
            '</button>' +
            '<button class="renewal-btn success" id="renewalConfirmBtn" ' +
                'onclick="userConfirmPaid()">' +
                '<i class="fas fa-check"></i> Tôi đã thanh toán' +
            '</button>' +
        '</div>';
}

window.copyText = function(text, btn) {
    if (navigator.clipboard) {
        navigator.clipboard.writeText(text).then(function() {
            btn.classList.add('copied');
            btn.innerHTML = '<i class="fas fa-check"></i>';
            setTimeout(function() {
                btn.classList.remove('copied');
                btn.innerHTML = '<i class="fas fa-copy"></i>';
            }, 1500);
        });
    } else {
        var ta = document.createElement('textarea');
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        btn.classList.add('copied');
        btn.innerHTML = '<i class="fas fa-check"></i>';
        setTimeout(function() {
            btn.classList.remove('copied');
            btn.innerHTML = '<i class="fas fa-copy"></i>';
        }, 1500);
    }
};

/* ─── USER BẤM "ĐÃ THANH TOÁN" ─────────────────────────────────── */
window.userConfirmPaid = async function() {
    if (!renewalCurrentReq) return;
    var btn = $('renewalConfirmBtn');
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Đang gửi...';

    try {
        await db.collection('renewal_requests').doc(renewalCurrentReq.id).update({
            userConfirmedAt: firebase.firestore.FieldValue.serverTimestamp(),
            status: 'user_paid'
        });
        renderPendingConfirm();
    } catch(e) {
        alert('❌ Lỗi: ' + e.message);
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-check"></i> Tôi đã thanh toán';
    }
};

function renderPendingConfirm() {
    var zaloUrl = RENEWAL_SUPPORT_ZALO
        ? 'https://zalo.me/' + RENEWAL_SUPPORT_ZALO.replace(/\D/g, '')
        : '#';

    $('renewalBody').innerHTML =
        '<div class="renewal-success">' +
            '<div class="icon" style="background:linear-gradient(135deg,#f59e0b,#d97706);' +
                'box-shadow:0 8px 24px rgba(245,158,11,.4)">' +
                '<i class="fas fa-hourglass-half"></i>' +
            '</div>' +
            '<h3>Đang chờ xác nhận</h3>' +
            '<p>Hệ thống đang kiểm tra giao dịch của bạn.<br>' +
                'Vui lòng đợi <b>1-5 phút</b>.</p>' +
            '<div class="info-box">' +
                '<i class="fas fa-info-circle" style="color:var(--primary)"></i>' +
                '<div>Nếu sau <b>10 phút</b> chưa được gia hạn, ' +
                'liên hệ Zalo hỗ trợ kèm mã: <b>' +
                escapeHtml(renewalCurrentReq.code) + '</b></div>' +
            '</div>' +
            '<div style="display:flex;gap:.5rem;flex-wrap:wrap;justify-content:center;margin-top:.5rem">' +
                '<a class="renewal-btn" href="' + zaloUrl +
                    '" target="_blank" rel="noopener">' +
                    '<i class="fas fa-comment-dots"></i> Liên hệ Zalo' +
                '</a>' +
                '<button class="renewal-btn primary" onclick="closeRenewalModal()">' +
                    '<i class="fas fa-check"></i> Đóng' +
                '</button>' +
            '</div>' +
        '</div>';
}

/* ─── LISTENER: LẮNG NGHE XÁC NHẬN ────────────────────────────── */
function listenRenewalRequest(reqId) {
    if (renewalListener) {
        try { renewalListener(); } catch(e) {}
    }
    renewalListener = db.collection('renewal_requests').doc(reqId)
        .onSnapshot(function(doc) {
            if (!doc.exists) return;
            var data = doc.data();
            if (data.status === 'confirmed') {
                showRenewalSuccess(data);
                if (renewalListener) {
                    try { renewalListener(); } catch(e) {}
                    renewalListener = null;
                }
                refreshCurrentUser();
            }
        });
}

function showRenewalSuccess(data) {
    var newExpiry = '';
    if (data.newExpiresAt) {
        try {
            var d = data.newExpiresAt.toDate
                ? data.newExpiresAt.toDate()
                : new Date(data.newExpiresAt.seconds * 1000);
            newExpiry = d.toLocaleDateString('vi-VN');
        } catch(e) {}
    }

    $('renewalBody').innerHTML =
        '<div class="renewal-success">' +
            '<div class="icon"><i class="fas fa-check"></i></div>' +
            '<h3>🎉 Gia hạn thành công!</h3>' +
            '<p>Tài khoản của bạn đã được gia hạn thêm <b>' +
                data.days + ' ngày</b>.</p>' +
            (newExpiry ? '<div class="info-box">' +
                '<i class="fas fa-calendar-check" style="color:var(--success)"></i>' +
                '<div>Hạn mới: <b>' + newExpiry + '</b></div>' +
            '</div>' : '') +
            '<button class="renewal-btn primary" ' +
                'onclick="closeRenewalModal();location.reload()" ' +
                'style="margin-top:.5rem">' +
                '<i class="fas fa-check"></i> Hoàn tất' +
            '</button>' +
        '</div>';
}

window.cancelRenewal = async function() {
    if (renewalCurrentReq) {
        try {
            await db.collection('renewal_requests').doc(renewalCurrentReq.id).update({
                status: 'cancelled',
                cancelledAt: firebase.firestore.FieldValue.serverTimestamp()
            });
        } catch(e) {}
    }
    closeRenewalModal();
};

/* ─── HELPERS ─────────────────────────────────────────────────── */
function generateTransferCode(email) {
    var rand = Math.random().toString(36).substring(2, 6).toUpperCase();
    var timestamp = Date.now().toString(36).slice(-4).toUpperCase();
    return 'HN' + timestamp + rand;
}

function formatMoney(n) {
    return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

async function refreshCurrentUser() {
    if (!currentUser) return;
    try {
        var doc = await db.collection('allowed_users').doc(currentUser.email).get();
        if (!doc.exists) return;
        var data = doc.data();
        currentUser.expiresAt = data.expiresAt || null;
        currentUser.name = data.name || currentUser.name;

        try { localStorage.removeItem('user_cache_' + currentUser.email); } catch(e) {}
        try {
            localStorage.setItem('user_cache_' + currentUser.email, JSON.stringify({
                data: currentUser,
                expires: Date.now() + 12 * 60 * 60 * 1000
            }));
        } catch(e) {}

        if (typeof applyUserUI === 'function') applyUserUI();
    } catch(e) {
        console.error('refreshCurrentUser error:', e);
    }
}

/* ─── ADMIN: LOAD + DUYỆT YÊU CẦU ─────────────────────────────── */
function loadRenewals() {
    if (typeof isHiddenAdmin === 'function' && isHiddenAdmin()) {
        var t = $('renewalsTitle');
        if (t) t.style.display = 'none';
        var l = $('renewalsList');
        if (l) l.style.display = 'none';
        return;
    }
    var titleEl = $('renewalsTitle');
    if (titleEl) titleEl.style.display = 'flex';
    var listEl = $('renewalsList');
    if (listEl) listEl.style.display = 'block';

    if (!listEl) return;

    db.collection('renewal_requests')
        .orderBy('createdAt', 'desc')
        .limit(100)
        .get()
        .then(function(snapshot) {
            var items = [];
            snapshot.forEach(function(doc) {
                var d = doc.data();
                if (d.status === 'pending' || d.status === 'user_paid') {
                    items.push(Object.assign({ _id: doc.id }, d));
                }
            });

            var pendingBadge = $('pendingRenewalsBadge');
            if (pendingBadge) {
                pendingBadge.textContent = items.length;
                pendingBadge.style.display = items.length > 0 ? 'inline-block' : 'none';
            }

            if (items.length === 0) {
                listEl.innerHTML = '<div class="no-data" style="padding:1rem;font-size:.8rem">' +
                    'Không có yêu cầu nào</div>';
                return;
            }

            var html = '';
            items.forEach(function(d) {
                var created = d.createdAt ? d.createdAt.toDate() : new Date();
                var timeStr = formatTimeDiff(Date.now() - created.getTime());
                var statusCls = d.status === 'user_paid' ? 'user_paid' : 'pending';
                var statusText = d.status === 'user_paid'
                    ? '⏳ Chờ xác nhận'
                    : '⏱ Chờ CK';

                html +=
                '<div class="renewal-admin-row">' +
                    '<div class="rar-head">' +
                        '<div class="u-info">' +
                            '<div class="rar-email">' + escapeHtml(d.name || d.email) + '</div>' +
                            '<div class="u-email">' + escapeHtml(d.email) + '</div>' +
                            '<div class="rar-sub">' +
                                '<i class="fas fa-clock"></i> ' + timeStr +
                            '</div>' +
                        '</div>' +
                        '<div class="rar-pkg">' +
                            '<div class="rar-amount">' + formatMoney(d.amount) + 'đ</div>' +
                            '<div class="rar-pkg-label">' +
                                escapeHtml(d.packageLabel || d.package) +
                                ' · ' + d.days + ' ngày' +
                            '</div>' +
                        '</div>' +
                    '</div>' +
                    '<div class="rar-foot">' +
                        '<div>' +
                            'Mã: <span class="rar-code">' +
                                escapeHtml(d.transferCode) + '</span>' +
                        '</div>' +
                        '<div class="rar-actions">' +
                            '<span class="rar-status ' + statusCls + '">' +
                                statusText + '</span>' +
                            '<button class="btn primary" ' +
                                'style="padding:.4rem .75rem;font-size:.75rem" ' +
                                'onclick="approveRenewal(\'' +
                                    escapeJs(d._id) + '\')">' +
                                '<i class="fas fa-check"></i> Xác nhận' +
                            '</button>' +
                            '<button class="btn" ' +
                                'style="padding:.4rem .65rem;font-size:.75rem" ' +
                                'onclick="rejectRenewal(\'' +
                                    escapeJs(d._id) + '\')">' +
                                '<i class="fas fa-times"></i>' +
                            '</button>' +
                        '</div>' +
                    '</div>' +
                '</div>';
            });
            listEl.innerHTML = html;
        })
        .catch(function(err) {
            listEl.innerHTML = '<div class="no-data" ' +
                'style="padding:1rem;font-size:.8rem;color:#dc2626">' +
                'Lỗi: ' + err.message + '</div>';
        });
}

window.approveRenewal = async function(reqId) {
    if (!confirm('Xác nhận đã nhận tiền và gia hạn tài khoản này?')) return;

    try {
        var reqDoc = await db.collection('renewal_requests').doc(reqId).get();
        if (!reqDoc.exists) {
            alert('Không tìm thấy yêu cầu!');
            return;
        }
        var req = reqDoc.data();

        var userDoc = await db.collection('allowed_users').doc(req.email).get();
        var currentExpiry = null;
        if (userDoc.exists) {
            var ud = userDoc.data();
            if (ud.expiresAt) {
                currentExpiry = ud.expiresAt.toDate
                    ? ud.expiresAt.toDate()
                    : new Date(ud.expiresAt.seconds * 1000);
            }
        }

        var now = new Date();
        var baseDate = (currentExpiry && currentExpiry > now) ? currentExpiry : now;
        var newExpiry = new Date(baseDate.getTime() + req.days * 24 * 60 * 60 * 1000);

        await db.collection('allowed_users').doc(req.email).update({
            expiresAt: firebase.firestore.Timestamp.fromDate(newExpiry),
            isTrial: false,
            lastRenewalAt: firebase.firestore.FieldValue.serverTimestamp()
        });

        await db.collection('renewal_requests').doc(reqId).update({
            status: 'confirmed',
            confirmedAt: firebase.firestore.FieldValue.serverTimestamp(),
            confirmedBy: currentUser.email,
            newExpiresAt: firebase.firestore.Timestamp.fromDate(newExpiry)
        });

        try { localStorage.removeItem('user_cache_' + req.email); } catch(e) {}

        alert('✅ Đã gia hạn ' + req.days + ' ngày cho ' + req.email +
              '\nHạn mới: ' + newExpiry.toLocaleDateString('vi-VN'));

        loadRenewals();
        if (typeof loadUsers === 'function') loadUsers(true);
    } catch(e) {
        alert('❌ Lỗi: ' + e.message);
    }
};

window.rejectRenewal = async function(reqId) {
    var reason = prompt('Lý do từ chối (tùy chọn):', '');
    if (reason === null) return;

    try {
        await db.collection('renewal_requests').doc(reqId).update({
            status: 'rejected',
            rejectedAt: firebase.firestore.FieldValue.serverTimestamp(),
            rejectedBy: currentUser.email,
            rejectReason: reason || ''
        });
        loadRenewals();
    } catch(e) {
        alert('❌ Lỗi: ' + e.message);
    }
};

/* ─── INIT ─────────────────────────────────────────────────────── */
function initRenewalUI() {
    var closeBtn = $('renewalClose');
    if (closeBtn) closeBtn.addEventListener('click', closeRenewalModal);

    var modal = $('renewalModal');
    if (modal) {
        modal.addEventListener('click', function(e) {
            if (e.target === this) closeRenewalModal();
        });
    }

    var dropdownBtn = $('dropdownRenewBtn');
    if (dropdownBtn) {
        dropdownBtn.addEventListener('click', function(e) {
            e.preventDefault();
            openRenewalModal();
        });
    }

    var bannerBtn = $('expiryRenewBtn');
    if (bannerBtn) {
        bannerBtn.addEventListener('click', function(e) {
            e.preventDefault();
            openRenewalModal();
        });
    }
}
"""
