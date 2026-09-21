# -*- coding: utf-8 -*-
"""
Template cho tài khoản: login Firebase, admin panel, user management,
expiry, import/export Excel, super admin / admin thường.
KHÔNG CẦN SỬA khi đổi cấu trúc Excel hay giao diện học.
"""


def build_accounts_css():
    """CSS: user menu, dropdown, admin panel, import, edit modals."""
    return r"""
/* ============ USER MENU ============ */
.user-menu{position:relative}
.user-avatar{
    width:38px;height:38px;border-radius:50%;border:2px solid var(--border);
    cursor:pointer;object-fit:cover;transition:.15s;display:block;
}
.user-avatar:hover{border-color:var(--primary);transform:scale(1.05)}
.user-dropdown{
    position:absolute;top:calc(100% + .5rem);right:0;
    background:var(--surface);border:1px solid var(--border);
    border-radius:var(--radius);box-shadow:0 10px 30px rgba(0,0,0,.15);
    padding:.5rem;min-width:290px;display:none;z-index:200;
}
.user-dropdown.show{display:block}
.user-info{padding:.75rem;border-bottom:1px solid var(--border);margin-bottom:.5rem}
.user-info .name{font-weight:700;font-size:.9rem;color:var(--text);margin-bottom:.2rem}
.user-info .email{font-size:.75rem;color:var(--text-3);word-break:break-all}
.user-info .role{
    display:inline-block;margin-top:.4rem;padding:.15rem .5rem;
    background:var(--primary-light);color:var(--primary-dark);
    border-radius:50px;font-size:.68rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.3px;
}
.user-info .role.admin{background:var(--amber-light);color:#92400e}
.dropdown-item{
    display:flex;align-items:center;gap:.5rem;width:100%;
    padding:.65rem .75rem;border:none;border-radius:var(--radius-sm);
    background:transparent;color:var(--text);font-size:.85rem;font-weight:600;
    cursor:pointer;transition:.15s;font-family:inherit;text-align:left;
}
.dropdown-item:hover{background:var(--surface-2)}
.dropdown-item.danger{color:var(--danger)}
.dropdown-item.danger:hover{background:var(--danger-light)}

.user-details{
    padding:.6rem .75rem .75rem .75rem;
    border-bottom:1px solid var(--border);margin-bottom:.5rem;
    display:flex;flex-direction:column;gap:.6rem;
}
.detail-row{display:flex;align-items:flex-start;gap:.65rem;}
.detail-icon{
    width:36px;height:36px;border-radius:10px;
    display:flex;align-items:center;justify-content:center;
    font-size:.95rem;flex-shrink:0;
    background:var(--surface-2);color:var(--text-2);transition:.2s;
}
.detail-icon.ok{background:rgba(22,163,74,.12);color:var(--success);}
.detail-icon.warn{background:rgba(245,158,11,.15);color:#d97706;}
.detail-icon.urgent{background:rgba(220,38,38,.15);color:var(--danger);}
.detail-icon.permanent{background:var(--primary-light);color:var(--primary-dark);}
.detail-icon.expired{background:rgba(220,38,38,.25);color:var(--danger);}
[data-theme="dark"] .detail-icon.ok{background:rgba(22,163,74,.25);color:#4ade80;}
[data-theme="dark"] .detail-icon.warn{background:rgba(245,158,11,.25);color:#fcd34d;}
[data-theme="dark"] .detail-icon.urgent{background:rgba(220,38,38,.3);color:#fca5a5;}
[data-theme="dark"] .detail-icon.permanent{background:rgba(59,130,246,.25);color:#93c5fd;}
[data-theme="dark"] .detail-icon.expired{background:rgba(220,38,38,.35);color:#fca5a5;}

.detail-content{flex:1;min-width:0;}
.detail-label{font-size:.65rem;font-weight:700;text-transform:uppercase;letter-spacing:.3px;color:var(--text-3);margin-bottom:.15rem;}
.detail-value{font-size:.85rem;font-weight:700;color:var(--text);line-height:1.3;word-break:break-word;}
.detail-value.ok{color:var(--success);}
.detail-value.warn{color:#d97706;}
[data-theme="dark"] .detail-value.warn{color:#fcd34d;}
.detail-value.urgent{color:var(--danger);}
.detail-value.permanent{color:var(--primary-dark);}
[data-theme="dark"] .detail-value.permanent{color:#93c5fd;}
.detail-value.expired{color:var(--danger);text-decoration:line-through;}
.detail-sub{font-size:.7rem;color:var(--text-3);margin-top:.15rem;line-height:1.35;}
.detail-sub b{color:var(--text-2);font-weight:700;}
.detail-progress{margin-top:.15rem;}
.progress-track{width:100%;height:6px;background:var(--surface-2);border-radius:50px;overflow:hidden;border:1px solid var(--border);}
.progress-bar{height:100%;border-radius:50px;transition:width .4s ease, background .3s ease;background:linear-gradient(90deg, #16a34a, #22c55e);}
.progress-bar.ok{background:linear-gradient(90deg, #16a34a, #22c55e);}
.progress-bar.warn{background:linear-gradient(90deg, #f59e0b, #fbbf24);}
.progress-bar.urgent{background:linear-gradient(90deg, #dc2626, #ef4444);}
.progress-bar.permanent{background:linear-gradient(90deg, #2563eb, #3b82f6);}

.btn-login-header{
    display:flex;align-items:center;gap:.4rem;
    padding:.55rem 1rem;border-radius:50px;
    background:var(--primary);color:#fff;border:none;
    font-size:.85rem;font-weight:700;cursor:pointer;
    transition:.15s;font-family:inherit;
    box-shadow:0 4px 12px rgba(37,99,235,.3);white-space:nowrap;
}
.btn-login-header:hover,.btn-login-header:active{background:var(--primary-dark);transform:translateY(-1px)}

/* ============ LOGIN MODAL ============ */
.login-modal{
    position:fixed;inset:0;background:rgba(15,23,42,.8);
    backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);
    z-index:3000;display:none;align-items:center;justify-content:center;
    padding:1.5rem;animation:fadeIn .2s;
}
.login-modal.show{display:flex}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
.login-box{
    background:#fff;border-radius:20px;padding:2.5rem 2rem;
    max-width:440px;width:100%;
    box-shadow:0 20px 60px rgba(0,0,0,.3);
    text-align:center;position:relative;
    animation:slideUp .3s cubic-bezier(.34,1.56,.64,1);
}
@keyframes slideUp{
    from{transform:translateY(30px) scale(.95);opacity:0}
    to{transform:translateY(0) scale(1);opacity:1}
}
.login-close{
    position:absolute;top:12px;right:12px;
    width:34px;height:34px;border-radius:50%;border:none;
    background:#f1f5f9;color:#475569;cursor:pointer;font-size:1rem;
    display:flex;align-items:center;justify-content:center;transition:.15s;
}
.login-close:hover{background:#fee2e2;color:#dc2626}
.login-logo{
    width:70px;height:70px;background:linear-gradient(135deg,#2563eb,#7c3aed);
    border-radius:20px;display:flex;align-items:center;justify-content:center;
    color:#fff;font-size:2rem;margin:0 auto 1.5rem;
    box-shadow:0 8px 20px rgba(37,99,235,.35);
}
.login-box h2{font-size:1.4rem;color:#0f172a;margin-bottom:.5rem;font-weight:700}
.login-box p{color:#64748b;font-size:.9rem;margin-bottom:2rem;line-height:1.5}
.btn-google{
    display:flex;align-items:center;justify-content:center;gap:.75rem;
    width:100%;padding:.9rem 1.5rem;border-radius:50px;
    border:2px solid #e2e8f0;background:#fff;color:#0f172a;
    font-size:1rem;font-weight:600;cursor:pointer;transition:.15s;font-family:inherit;
}
.btn-google:hover{border-color:#2563eb;background:#f0f7ff;transform:translateY(-1px);box-shadow:0 4px 12px rgba(37,99,235,.15)}
.btn-google img{width:22px;height:22px}
.login-error{
    background:#fee2e2;color:#dc2626;padding:.85rem 1rem;
    border-radius:10px;font-size:.85rem;margin-top:1rem;
    display:none;text-align:left;line-height:1.4;
}
.login-error.show{display:block}
.login-footer{margin-top:1.5rem;padding-top:1.5rem;border-top:1px solid #e2e8f0;font-size:.78rem;color:#94a3b8;line-height:1.5}

/* ============ ADMIN MODAL ============ */
.admin-modal{
    position:fixed;inset:0;background:rgba(15,23,42,.75);
    backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);
    z-index:3000;display:none;align-items:center;justify-content:center;
    padding:1rem;animation:fadeIn .2s;
}
.admin-modal.show{display:flex}
.admin-box{
    background:var(--surface);border-radius:20px;width:100%;max-width:900px;
    max-height:calc(100vh - 2rem);overflow:hidden;
    box-shadow:0 20px 60px rgba(0,0,0,.3);
    display:flex;flex-direction:column;
}
.admin-header{
    padding:1.25rem 1.5rem;border-bottom:1px solid var(--border);
    display:flex;align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap;
}
.admin-header h2{font-size:1.15rem;color:var(--text);display:flex;align-items:center;gap:.5rem;font-weight:700}
.admin-header h2 i{color:var(--amber)}
.admin-header-actions{display:flex;gap:.5rem;align-items:center;flex-wrap:wrap}
.admin-body{padding:1.25rem 1.5rem;overflow-y:auto;flex:1}
.admin-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:.75rem;margin-bottom:1.25rem;}
.stat-card{background:var(--surface-2);border:1px solid var(--border);border-radius:var(--radius);padding:1rem;text-align:center;}
.stat-card .num{font-size:1.8rem;font-weight:800;color:var(--primary);line-height:1;margin-bottom:.3rem}
.stat-card .label{font-size:.75rem;color:var(--text-3);text-transform:uppercase;letter-spacing:.3px;font-weight:600}
.admin-section-title{font-size:.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text-3);margin-bottom:.6rem;display:flex;align-items:center;justify-content:space-between;}
.admin-section-title i{margin-right:.3rem;}
.btn-add{padding:.45rem .85rem;border-radius:8px;border:none;background:var(--primary);color:#fff;font-size:.78rem;font-weight:600;cursor:pointer;display:inline-flex;align-items:center;gap:.35rem;transition:.15s;font-family:inherit;}
.btn-add:hover{background:var(--primary-dark)}
.user-list{display:flex;flex-direction:column;gap:.5rem}
.user-row{display:flex;align-items:center;gap:.75rem;padding:.75rem;background:var(--surface-2);border:1px solid var(--border);border-radius:var(--radius-sm);transition:.15s;}
.user-row:hover{border-color:var(--primary)}
.user-row .u-info{flex:1;min-width:0}
.user-row .u-name{font-weight:700;font-size:.88rem;color:var(--text);margin-bottom:.15rem}
.user-row .u-email{font-size:.75rem;color:var(--text-3);word-break:break-all}
.user-row .u-role{padding:.15rem .5rem;border-radius:50px;font-size:.65rem;font-weight:700;text-transform:uppercase;letter-spacing:.3px;white-space:nowrap;}
.user-row .u-role.admin{background:var(--amber-light);color:#92400e}
.user-row .u-role.user{background:var(--primary-light);color:var(--primary-dark)}
.user-row .u-role.super{background:linear-gradient(135deg, #f59e0b, #d97706);color:#fff;box-shadow:0 2px 6px rgba(245,158,11,.4);}
.user-row .u-actions{display:flex;gap:.3rem}
.u-btn{width:32px;height:32px;border-radius:8px;border:1px solid var(--border);background:var(--surface);color:var(--text-2);cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:.8rem;transition:.15s;}
.u-btn:hover{background:var(--surface-2);color:var(--primary);border-color:var(--primary)}
.u-btn.danger:hover{background:var(--danger-light);color:var(--danger);border-color:var(--danger)}
.u-btn.expiry{background:rgba(245,158,11,.1);color:#d97706;border-color:rgba(245,158,11,.4);}
.u-btn.expiry:hover{background:var(--amber);color:#fff;border-color:var(--amber);transform:scale(1.08);}
.u-btn.expiry:active{transform:scale(.95);}
[data-theme="dark"] .u-btn.expiry{background:rgba(245,158,11,.2);color:#fcd34d;border-color:rgba(245,158,11,.5);}
[data-theme="dark"] .u-btn.expiry:hover{background:var(--amber);color:#fff;}
.u-btn.expiry.urgent{background:rgba(220,38,38,.15);color:var(--danger);border-color:rgba(220,38,38,.5);animation:expiryPulse 2s infinite;}
.u-btn.expiry.urgent:hover{background:var(--danger);color:#fff;border-color:var(--danger);}
[data-theme="dark"] .u-btn.expiry.urgent{background:rgba(220,38,38,.3);color:#fca5a5;border-color:rgba(220,38,38,.6);}
@keyframes expiryPulse{
    0%,100%{box-shadow:0 0 0 0 rgba(220,38,38,.4);}
    50%{box-shadow:0 0 0 6px rgba(220,38,38,0);}
}
.u-btn:disabled{opacity:.35;cursor:not-allowed}
.u-btn:disabled:hover{background:var(--surface);color:var(--text-2);border-color:var(--border);transform:none;}
.u-btn.danger:disabled:hover{background:var(--surface);color:var(--text-2);border-color:var(--border);}
.u-btn.expiry:disabled:hover{background:var(--surface);color:var(--text-2);border-color:var(--border);transform:none;}

.admin-close{width:34px;height:34px;border-radius:8px;border:1px solid var(--border);background:var(--surface);color:var(--text-2);cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:.9rem;transition:.15s;}
.admin-close:hover{background:var(--danger-light);color:var(--danger);border-color:var(--danger)}
.add-user-form{background:var(--surface-2);border:1px solid var(--border);border-radius:var(--radius-sm);padding:1rem;margin-bottom:1rem;display:none;}
.add-user-form.show{display:block}
.add-user-form h3{font-size:.85rem;color:var(--text);margin-bottom:.75rem;font-weight:700}
.form-group{margin-bottom:.75rem}
.form-group label{display:block;font-size:.75rem;font-weight:600;color:var(--text-2);margin-bottom:.3rem}
.form-group input,.form-group select{width:100%;padding:.6rem .85rem;border-radius:8px;border:1.5px solid var(--border);background:var(--surface);color:var(--text);font-size:.85rem;outline:none;transition:.15s;font-family:inherit;}
.form-group input:focus,.form-group select:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(37,99,235,.15)}
.form-actions{display:flex;gap:.5rem;justify-content:flex-end;margin-top:.75rem}
.btn{padding:.55rem 1rem;border-radius:8px;border:1.5px solid var(--border);background:var(--surface);color:var(--text);font-size:.82rem;font-weight:600;cursor:pointer;transition:.15s;font-family:inherit;display:inline-flex;align-items:center;gap:.35rem;}
.btn:hover{background:var(--surface-2)}
.btn:disabled{opacity:.5;cursor:not-allowed;}
.btn.primary{background:var(--primary);color:#fff;border-color:var(--primary)}
.btn.primary:hover{background:var(--primary-dark)}
.logs-list{max-height:200px;overflow-y:auto;background:var(--surface-2);border:1px solid var(--border);border-radius:var(--radius-sm);padding:.5rem;}
.log-item{display:flex;gap:.5rem;padding:.4rem .5rem;font-size:.75rem;color:var(--text-2);border-bottom:1px solid var(--border);}
.log-item:last-child{border-bottom:none}
.log-item .log-time{color:var(--text-3);flex-shrink:0;font-family:monospace;font-size:.7rem}
.log-item .log-msg{flex:1;word-break:break-word}
.u-last-login{font-size:.68rem;color:var(--text-3);display:flex;align-items:center;gap:.25rem;margin-top:.2rem;}
.u-last-login.active{color:var(--success);}
.u-last-login.recent{color:var(--primary);}
.u-last-login i{font-size:.65rem;}
.u-expiry{font-size:.68rem;font-weight:600;display:inline-flex;align-items:center;gap:.25rem;margin-top:.2rem;padding:.15rem .45rem;border-radius:50px;cursor:pointer;transition:.15s;}
.u-expiry:hover{opacity:.8;}
.u-expiry i{font-size:.6rem;}
.u-expiry.permanent{background:rgba(148,163,184,.15);color:var(--text-3);}
.u-expiry.ok{background:rgba(22,163,74,.12);color:var(--success);}
.u-expiry.warn{background:rgba(245,158,11,.15);color:#92400e;}
[data-theme="dark"] .u-expiry.warn{color:#fcd34d;}
.u-expiry.urgent{background:rgba(220,38,38,.15);color:var(--danger);}
.u-expiry.expired{background:rgba(220,38,38,.25);color:#fff;text-decoration:line-through;}

/* ============ EDIT MODALS ============ */
.edit-modal{position:fixed;inset:0;background:rgba(15,23,42,.85);backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);z-index:4000;display:none;align-items:center;justify-content:center;padding:1rem;animation:fadeIn .2s;}
.edit-modal.show{display:flex}
.edit-box{background:var(--surface);border-radius:20px;width:100%;max-width:420px;box-shadow:0 20px 60px rgba(0,0,0,.4);padding:1.5rem;position:relative;animation:slideUp .3s cubic-bezier(.34,1.56,.64,1);}
.edit-box h2{font-size:1.1rem;color:var(--text);font-weight:700;display:flex;align-items:center;gap:.5rem;margin-bottom:1.25rem;}
.edit-box h2 i{color:var(--primary);}
.edit-close{position:absolute;top:12px;right:12px;width:32px;height:32px;border-radius:8px;border:1px solid var(--border);background:var(--surface);color:var(--text-2);cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:.85rem;transition:.15s;}
.edit-close:hover{background:var(--danger-light);color:var(--danger);border-color:var(--danger)}
.edit-user-info{padding:.75rem;background:var(--surface-2);border-radius:10px;margin-bottom:1rem;border:1px solid var(--border);}
.edit-user-info .eu-name{font-weight:700;font-size:.9rem;color:var(--text);margin-bottom:.2rem}
.edit-user-info .eu-email{font-size:.75rem;color:var(--text-3);word-break:break-all}
.quick-expiry-btns{display:grid;grid-template-columns:repeat(3,1fr);gap:.4rem;margin-bottom:1rem;}
.quick-expiry-btn{padding:.5rem .4rem;border-radius:8px;border:1.5px solid var(--border);background:var(--surface);color:var(--text-2);font-size:.72rem;font-weight:600;cursor:pointer;transition:.15s;font-family:inherit;display:flex;flex-direction:column;align-items:center;gap:.2rem;white-space:nowrap;}
.quick-expiry-btn i{font-size:.85rem;}
.quick-expiry-btn:hover{background:var(--primary-light);border-color:var(--primary);color:var(--primary-dark);transform:translateY(-1px);}
.quick-expiry-btn.danger:hover{background:var(--danger-light);border-color:var(--danger);color:var(--danger);}

/* ============ IMPORT MODAL ============ */
.import-modal{position:fixed;inset:0;background:rgba(15,23,42,.85);backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);z-index:3500;display:none;align-items:center;justify-content:center;padding:1rem;animation:fadeIn .2s;}
.import-modal.show{display:flex}
.import-box{background:var(--surface);border-radius:20px;width:100%;max-width:900px;max-height:calc(100vh - 2rem);box-shadow:0 20px 60px rgba(0,0,0,.4);display:flex;flex-direction:column;overflow:hidden;}
.import-header{padding:1.25rem 1.5rem;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;}
.import-header h2{font-size:1.15rem;color:var(--text);display:flex;align-items:center;gap:.5rem;font-weight:700;}
.import-header h2 i{color:#16a34a;}
.import-body{padding:1.25rem 1.5rem;overflow-y:auto;flex:1;min-height:0;}
.import-summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:.6rem;margin-bottom:1rem;}
.import-stat{padding:.65rem .85rem;border-radius:10px;text-align:center;border:1px solid var(--border);background:var(--surface-2);}
.import-stat .num{font-size:1.5rem;font-weight:800;line-height:1;margin-bottom:.25rem;}
.import-stat .label{font-size:.7rem;color:var(--text-3);text-transform:uppercase;font-weight:600;}
.import-stat.ok .num{color:var(--success);}
.import-stat.update .num{color:var(--primary);}
.import-stat.warn .num{color:var(--amber);}
.import-stat.err .num{color:var(--danger);}
.import-preview-wrap{max-height:400px;overflow-y:auto;border:1px solid var(--border);border-radius:10px;background:var(--surface-2);}
.import-table{width:100%;border-collapse:collapse;font-size:.82rem;}
.import-table th{padding:.6rem .8rem;text-align:left;background:var(--surface);color:var(--text-2);font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.3px;border-bottom:2px solid var(--border);position:sticky;top:0;z-index:2;}
.import-table td{padding:.55rem .8rem;color:var(--text);border-bottom:1px solid var(--border);word-break:break-word;}
.import-table tr:last-child td{border-bottom:none;}
.import-table tr.row-error{background:rgba(220,38,38,.08);}
.import-table tr.row-warn{background:rgba(245,158,11,.08);}
.import-table tr.row-new{background:rgba(22,163,74,.05);}
.import-table tr.row-update{background:rgba(37,99,235,.05);}
.import-table .status-badge{display:inline-flex;align-items:center;gap:.3rem;padding:.2rem .5rem;border-radius:50px;font-size:.68rem;font-weight:700;white-space:nowrap;}
.import-table .status-badge.ok{background:rgba(22,163,74,.15);color:var(--success);}
.import-table .status-badge.update{background:rgba(37,99,235,.15);color:var(--primary);}
.import-table .status-badge.warn{background:rgba(245,158,11,.15);color:#92400e;}
.import-table .status-badge.err{background:rgba(220,38,38,.15);color:var(--danger);}
.import-table .role-badge{display:inline-block;padding:.15rem .5rem;border-radius:50px;font-size:.68rem;font-weight:700;text-transform:uppercase;}
.import-table .role-badge.admin{background:var(--amber-light);color:#92400e;}
.import-table .role-badge.user{background:var(--primary-light);color:var(--primary-dark);}
.import-options{display:flex;gap:1rem;margin-top:1rem;flex-wrap:wrap;}
.import-options label{display:flex;align-items:center;gap:.4rem;font-size:.82rem;font-weight:600;color:var(--text-2);cursor:pointer;user-select:none;}
.import-options input[type="checkbox"]{width:16px;height:16px;accent-color:var(--primary);cursor:pointer;}
.import-info{margin-top:1rem;padding:.65rem .85rem;border-radius:10px;background:var(--primary-light);color:var(--primary-dark);font-size:.78rem;line-height:1.6;display:flex;align-items:flex-start;gap:.5rem;}
[data-theme="dark"] .import-info{color:#93c5fd;}
.import-info i{margin-top:.15rem;flex-shrink:0;}
.import-info b{font-weight:800;}
.import-footer{padding:1rem 1.5rem;border-top:1px solid var(--border);display:flex;gap:.5rem;justify-content:flex-end;align-items:center;background:var(--surface);}

/* ============ EXPIRY BANNER ============ */
.expiry-banner{
    background:linear-gradient(135deg, #fef3c7, #fde68a);
    border:1.5px solid #f59e0b;border-radius:var(--radius);
    padding:.85rem 1.1rem;margin-bottom:1rem;
    display:flex;align-items:center;gap:.75rem;flex-wrap:wrap;
}
[data-theme="dark"] .expiry-banner{background:linear-gradient(135deg, rgba(245,158,11,.15), rgba(245,158,11,.25));}
.expiry-banner.urgent{background:linear-gradient(135deg, #fecaca, #fca5a5);border-color:#dc2626;animation:pulseUrgent 2s infinite;}
[data-theme="dark"] .expiry-banner.urgent{background:linear-gradient(135deg, rgba(220,38,38,.2), rgba(220,38,38,.3));}
@keyframes pulseUrgent{
    0%,100%{box-shadow:0 0 0 0 rgba(220,38,38,.4);}
    50%{box-shadow:0 0 0 8px rgba(220,38,38,0);}
}
.expiry-banner-icon{width:36px;height:36px;border-radius:50%;background:#f59e0b;color:#fff;display:flex;align-items:center;justify-content:center;font-size:1rem;flex-shrink:0;}
.expiry-banner.urgent .expiry-banner-icon{background:#dc2626;}
.expiry-banner-text{flex:1;min-width:200px}
.expiry-banner-text .title{font-weight:700;font-size:.9rem;color:#92400e;margin-bottom:.15rem;}
[data-theme="dark"] .expiry-banner-text .title{color:#fcd34d;}
.expiry-banner.urgent .expiry-banner-text .title{color:#991b1b;}
[data-theme="dark"] .expiry-banner.urgent .expiry-banner-text .title{color:#fca5a5;}
.expiry-banner-text .desc{font-size:.78rem;color:#78350f;line-height:1.5;}
[data-theme="dark"] .expiry-banner-text .desc{color:#fde68a;}
.expiry-banner-text b{color:#dc2626;}
.expiry-banner.urgent .expiry-banner-text b{color:#991b1b;}
.expiry-banner-btn{padding:.5rem .9rem;border-radius:50px;border:none;background:#f59e0b;color:#fff;text-decoration:none;font-size:.8rem;font-weight:700;cursor:pointer;transition:.15s;display:inline-flex;align-items:center;gap:.35rem;white-space:nowrap;}
.expiry-banner-btn:hover{background:#d97706;color:#fff;transform:translateY(-1px);}
.expiry-banner.urgent .expiry-banner-btn{background:#dc2626;}
.expiry-banner.urgent .expiry-banner-btn:hover{background:#b91c1c;}
"""


def build_accounts_html():
    """HTML cho login + user menu + admin + edit modals."""
    return r"""
<div class="login-modal" id="loginModal">
    <div class="login-box">
        <button class="login-close" id="loginClose"><i class="fas fa-times"></i></button>
        <div class="login-logo"><i class="fas fa-language"></i></div>
        <h2>Đăng nhập để mở khóa</h2>
        <p>Đăng nhập bằng Google để sử dụng <b>toàn bộ câu</b>, tất cả bộ lọc HSK1-6, chủ đề đầy đủ, luyện viết không giới hạn và nhiều tính năng khác.</p>
        <button class="btn-google" id="googleLoginBtn">
            <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google">
            Đăng nhập bằng Google
        </button>
        <div class="login-error" id="loginError"></div>
        <div class="login-footer">
            <i class="fas fa-shield-alt"></i> Chỉ tài khoản được cấp phép mới truy cập được.
        </div>
    </div>
</div>

<div class="edit-modal" id="changeNameModal">
    <div class="edit-box">
        <button class="edit-close" id="changeNameClose"><i class="fas fa-times"></i></button>
        <h2><i class="fas fa-user-edit"></i> Đổi tên hiển thị</h2>
        <div class="edit-user-info">
            <div class="eu-name" id="changeNameCurrent">-</div>
            <div class="eu-email" id="changeNameEmail">-</div>
        </div>
        <div class="form-group">
            <label>Tên mới</label>
            <input type="text" id="changeNameInput" placeholder="Nhập tên mới..." maxlength="50">
        </div>
        <div class="form-actions">
            <button class="btn" id="changeNameCancel">Hủy</button>
            <button class="btn primary" id="changeNameConfirm"><i class="fas fa-check"></i> Lưu</button>
        </div>
    </div>
</div>

<div class="edit-modal" id="editExpiryModal">
    <div class="edit-box">
        <button class="edit-close" id="editExpiryClose"><i class="fas fa-times"></i></button>
        <h2><i class="fas fa-calendar-edit"></i> Chỉnh hạn sử dụng</h2>
        <div class="edit-user-info">
            <div class="eu-name" id="editExpiryName">-</div>
            <div class="eu-email" id="editExpiryEmail">-</div>
        </div>
        <div class="quick-expiry-btns">
            <button class="quick-expiry-btn" onclick="setQuickExpiry(7)"><i class="fas fa-calendar-plus"></i> +7 ngày</button>
            <button class="quick-expiry-btn" onclick="setQuickExpiry(30)"><i class="fas fa-calendar-plus"></i> +30 ngày</button>
            <button class="quick-expiry-btn" onclick="setQuickExpiry(90)"><i class="fas fa-calendar-plus"></i> +90 ngày</button>
            <button class="quick-expiry-btn" onclick="setQuickExpiry(180)"><i class="fas fa-calendar-plus"></i> +6 tháng</button>
            <button class="quick-expiry-btn" onclick="setQuickExpiry(365)"><i class="fas fa-calendar-plus"></i> +1 năm</button>
            <button class="quick-expiry-btn danger" onclick="setQuickExpiryPermanent()"><i class="fas fa-infinity"></i> Vĩnh viễn</button>
        </div>
        <div class="form-group">
            <label>Hoặc chọn ngày cụ thể</label>
            <input type="date" id="editExpiryInput">
        </div>
        <div class="form-actions">
            <button class="btn" id="editExpiryCancel">Hủy</button>
            <button class="btn primary" id="editExpiryConfirm"><i class="fas fa-check"></i> Lưu</button>
        </div>
    </div>
</div>

<div class="import-modal" id="importModal">
    <div class="import-box">
        <div class="import-header">
            <h2><i class="fas fa-file-import"></i> Import danh sách user</h2>
            <button class="admin-close" id="importClose"><i class="fas fa-times"></i></button>
        </div>
        <div class="import-body">
            <div class="import-summary" id="importSummary"></div>
            <div class="import-preview-wrap">
                <table class="import-table">
                    <thead>
                        <tr>
                            <th style="width:45px">#</th>
                            <th>Email</th>
                            <th>Tên</th>
                            <th style="width:80px">Vai trò</th>
                            <th style="width:110px">Hạn dùng</th>
                            <th style="width:130px">Trạng thái</th>
                        </tr>
                    </thead>
                    <tbody id="importTableBody"></tbody>
                </table>
            </div>
            <div class="import-options">
                <label><input type="checkbox" id="importSkipDuplicates"> Bỏ qua user đã tồn tại (không update)</label>
                <label><input type="checkbox" id="importSkipInvalid" checked> Bỏ qua dòng không hợp lệ</label>
            </div>
            <div class="import-info">
                <i class="fas fa-info-circle"></i>
                <div>
                    File Excel cần có cột: <b>email</b> (bắt buộc), <b>name</b> (tùy chọn), <b>expiresAt</b> (tùy chọn).
                    <br>• <b>email</b>: bắt buộc, phải hợp lệ
                    <br>• <b>name</b>: nếu thiếu sẽ lấy phần trước @ của email
                    <br>• <b>expiresAt</b>: định dạng <b>YYYY-MM-DD</b>, để trống = vĩnh viễn
                    <br>• <b>Các cột khác</b> sẽ được <b>tự động bỏ qua</b>
                    <br><b>⚠️ Lưu ý:</b> Chỉ import <b>user</b>, KHÔNG import admin.
                </div>
            </div>
        </div>
        <div class="import-footer">
            <button class="btn" id="importCancelBtn">Hủy</button>
            <button class="btn primary" id="importConfirmBtn"><i class="fas fa-check"></i> Import <span id="importCount">0</span> user</button>
        </div>
    </div>
</div>

<div class="admin-modal" id="adminModal">
    <div class="admin-box">
        <div class="admin-header">
            <h2><i class="fas fa-shield-alt"></i> Quản lý tài khoản</h2>
            <div class="admin-header-actions">
                <button class="btn" id="exportExcelBtn" title="Xuất danh sách USER ra Excel"><i class="fas fa-file-export"></i> Export</button>
                <button class="btn" id="importExcelBtn" title="Import từ Excel (chỉ import user)"><i class="fas fa-file-import"></i> Import</button>
                <input type="file" id="importFileInput" accept=".xlsx,.xls,.csv" style="display:none">
                <button class="btn" id="refreshUsersBtn" title="Làm mới"><i class="fas fa-sync-alt"></i></button>
                <button class="admin-close" id="adminClose"><i class="fas fa-times"></i></button>
            </div>
        </div>
        <div class="admin-body">
            <div class="admin-stats" id="adminStats"></div>
            <div class="admin-section-title">
                <span><i class="fas fa-users"></i> Danh sách tài khoản (<span id="adminUserCount">0</span>)</span>
                <button class="btn-add" id="showAddUserBtn"><i class="fas fa-plus"></i> Thêm</button>
            </div>
            <div class="add-user-form" id="addUserForm">
                <h3>Thêm tài khoản mới</h3>
                <div class="form-group">
                    <label>Email Google</label>
                    <input type="email" id="newUserEmail" placeholder="user@gmail.com">
                </div>
                <div class="form-group">
                    <label>Tên hiển thị</label>
                    <input type="text" id="newUserName" placeholder="Nguyễn Văn A">
                </div>
                <div class="form-group">
                    <label>Vai trò</label>
                    <select id="newUserRole">
                        <option value="user">User (chỉ học)</option>
                        <option value="admin">Admin (quản trị)</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Hạn sử dụng (để trống = vĩnh viễn)</label>
                    <input type="date" id="newUserExpires">
                </div>
                <div class="form-actions">
                    <button class="btn" id="cancelAddUser">Hủy</button>
                    <button class="btn primary" id="confirmAddUser"><i class="fas fa-check"></i> Thêm</button>
                </div>
            </div>
            <div class="user-list" id="userList">
                <div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>
            </div>
            <div class="admin-section-title" style="margin-top:1.5rem" id="logsTitle">
                <span><i class="fas fa-history"></i> Lịch sử đăng nhập (gần đây)</span>
            </div>
            <div class="logs-list" id="logsList">
                <div class="no-data" style="padding:1rem;font-size:.8rem"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>
            </div>
        </div>
    </div>
</div>
"""


def build_accounts_js():
    """JS: Firebase auth, login, admin panel, user management, import/export."""
    return r"""
/* ============ AUTH ============ */
var currentUser = null;
var isDemo = true;
var auth, db;
var usersCache = [];
var lastLoginMap = {};
var importRows = [];
var editingEmail = null;
var editingExpiryEmail = null;
var appInitialized = false;

function isSuperAdmin() {
    if (!currentUser || currentUser.role !== 'admin') return false;
    var email = (currentUser.email || '').toLowerCase().trim();
    return email === SUPER_ADMIN.toLowerCase().trim();
}
function isHiddenAdmin() {
    if (!currentUser || currentUser.role !== 'admin') return false;
    return !isSuperAdmin();
}

try {
    firebase.initializeApp(FIREBASE_CONFIG);
    auth = firebase.auth();
    db = firebase.firestore();
    auth.onAuthStateChanged(handleAuthChange);
} catch(e) {
    console.error('Firebase init error:', e);
    enterDemoMode();
}

async function handleAuthChange(user) {
    if (!user) {
        currentUser = null; isDemo = true;
        applyUserUI(); enterDemoMode();
        return;
    }

    var email = (user.email || '').toLowerCase();
    var cacheKey = 'user_cache_' + email;
    var cached = null;
    try { cached = JSON.parse(localStorage.getItem(cacheKey) || 'null'); } catch(e) {}

    if (cached && cached.expires > Date.now() && cached.data) {
        if (!checkUserExpiration(cached.data)) {
            await auth.signOut();
            try { localStorage.removeItem(cacheKey); } catch(e) {}
            enterDemoMode();
            return;
        }
        currentUser = cached.data;
        isDemo = false;
        applyUserUI();
        logLogin(currentUser);
        if (!appInitialized) { initApp(); appInitialized = true; }
        else { if (typeof refreshApp === 'function') refreshApp(); }
        return;
    }

    try {
        var doc = await db.collection('allowed_users').doc(email).get();
        if (!doc.exists) {
            await auth.signOut();
            showLoginError('Tài khoản <b>' + email + '</b> chưa được cấp quyền.');
            enterDemoMode();
            return;
        }
        var data = doc.data();
        var userData = {
            email: email,
            name: data.name || user.displayName || email.split('@')[0],
            role: data.role || 'user',
            photo: user.photoURL || '',
            expiresAt: data.expiresAt || null
        };

        if (!checkUserExpiration(userData)) {
            await auth.signOut();
            try { localStorage.removeItem(cacheKey); } catch(e) {}
            enterDemoMode();
            return;
        }

        currentUser = userData;

        try {
            localStorage.setItem(cacheKey, JSON.stringify({
                data: currentUser,
                expires: Date.now() + 12 * 60 * 60 * 1000
            }));
        } catch(e) {}

        isDemo = false;
        applyUserUI();
        logLogin(currentUser);
        if (!appInitialized) { initApp(); appInitialized = true; }
        else { if (typeof refreshApp === 'function') refreshApp(); }
    } catch(e) {
        console.error('Auth check error:', e);
        isDemo = true; enterDemoMode();
    }
}

function checkUserExpiration(userData) {
    if (userData.role === 'admin') return true;
    if (!userData.expiresAt) return true;

    var expDate;
    try {
        var ea = userData.expiresAt;
        if (typeof ea.toDate === 'function') expDate = ea.toDate();
        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
        else expDate = new Date(ea);
    } catch(e) { return true; }

    if (expDate < new Date()) {
        var dateStr = expDate.toLocaleDateString('vi-VN');
        showLoginError('🔒 Tài khoản của bạn đã <b>hết hạn</b> vào ngày <b>' + dateStr + '</b>.<br><br>Vui lòng liên hệ Admin để gia hạn.');
        return false;
    }
    return true;
}

function getDaysRemaining(userData) {
    if (!userData || !userData.expiresAt) return null;
    if (userData.role === 'admin') return null;
    var expDate;
    try {
        var ea = userData.expiresAt;
        if (typeof ea.toDate === 'function') expDate = ea.toDate();
        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
        else expDate = new Date(ea);
    } catch(e) { return null; }
    return Math.ceil((expDate.getTime() - Date.now()) / (24 * 60 * 60 * 1000));
}

function enterDemoMode() {
    currentUser = null; isDemo = true;
    applyUserUI();
    if (!appInitialized) { initApp(); appInitialized = true; }
    else { if (typeof refreshApp === 'function') refreshApp(); }
    $('loadingScreen').classList.add('hidden');
    $('stickyTop').style.display = 'block';
    $('fabGroup').style.display = 'flex';
    $('mainContent').style.display = 'block';
}

function applyUserUI() {
    var demoBadge = $('demoBadge');
    var headerLoginBtn = $('headerLoginBtn');
    var userMenu = $('userMenu');

    var expiryBanner = $('expiryBanner');
    if (expiryBanner) {
        if (!isDemo && currentUser && currentUser.role !== 'admin') {
            var daysLeft = getDaysRemaining(currentUser);
            if (daysLeft !== null && daysLeft <= 7) {
                expiryBanner.style.display = 'flex';
                $('expiryDaysText').textContent = daysLeft > 0 ? daysLeft : 0;
                var expDate;
                try {
                    var ea = currentUser.expiresAt;
                    if (typeof ea.toDate === 'function') expDate = ea.toDate();
                    else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
                    else expDate = new Date(ea);
                } catch(e) {}
                if (expDate) $('expiryDateText').textContent = expDate.toLocaleDateString('vi-VN');
                expiryBanner.classList.toggle('urgent', daysLeft <= 3);
                var contactBtn = $('expiryContactBtn');
                if (ZALO_PHONE) {
                    var phone = ZALO_PHONE.replace(/\D/g, '');
                    contactBtn.href = 'https://zalo.me/' + phone;
                } else {
                    contactBtn.href = '#';
                    contactBtn.onclick = function(e) { e.preventDefault(); alert('Liên hệ Admin để gia hạn!'); };
                }
            } else {
                expiryBanner.style.display = 'none';
            }
        } else {
            expiryBanner.style.display = 'none';
        }
    }

    if (isDemo) {
        demoBadge.style.display = 'flex';
        headerLoginBtn.style.display = 'flex';
        userMenu.style.display = 'none';
        $('demoBanner').style.display = 'flex';
        var ud0 = $('userDetails');
        if (ud0) ud0.style.display = 'none';
    } else {
        demoBadge.style.display = 'none';
        headerLoginBtn.style.display = 'none';
        userMenu.style.display = 'block';
        $('demoBanner').style.display = 'none';

        $('userName').textContent = currentUser.name;
        $('userEmail').textContent = currentUser.email;
        $('userRole').textContent = currentUser.role;
        $('userRole').className = 'role' + (currentUser.role === 'admin' ? ' admin' : '');
        $('openAdminBtn').style.display = currentUser.role === 'admin' ? 'flex' : 'none';

        var avatar = $('userAvatar');
        if (currentUser.photo) avatar.src = currentUser.photo;
        else {
            avatar.src = 'data:image/svg+xml;utf8,' + encodeURIComponent(
                '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">' +
                '<rect fill="#2563eb" width="100" height="100"/>' +
                '<text x="50" y="65" font-size="45" fill="#fff" text-anchor="middle" font-family="sans-serif" font-weight="bold">' +
                currentUser.name.charAt(0).toUpperCase() + '</text></svg>'
            );
        }

        updateUserDetails();
    }

    var hskChip = $('hskChip');
    var subjectChip = $('subjectChip');
    if (isDemo) {
        hskChip.classList.add('demo-limited');
        subjectChip.classList.add('demo-limited');
    } else {
        hskChip.classList.remove('demo-limited');
        subjectChip.classList.remove('demo-limited');
    }

    var zaloBtn = $('zaloBtn');
    if (zaloBtn) {
        if (isDemo) zaloBtn.classList.remove('compact');
        else zaloBtn.classList.add('compact');
    }

    $('demoLimitText').textContent = DEMO_LIMIT;
    $('demoDailyText').textContent = DEMO_DAILY_LIMIT;
    var hskMaxEl1 = $('demoHskMaxText');
    var hskMaxEl2 = $('demoHskMaxText2');
    if (hskMaxEl1) hskMaxEl1.textContent = DEMO_HSK_MAX;
    if (hskMaxEl2) hskMaxEl2.textContent = DEMO_HSK_MAX;
    if (typeof updateDemoRemaining === 'function') updateDemoRemaining();

    var toggleFocusBtn = $('toggleFocusBtn');
    if (toggleFocusBtn) {
        toggleFocusBtn.style.display = isDemo ? 'none' : 'flex';
    }

    if (isDemo) {
        document.body.classList.remove('hide-floating');
    }
}

function updateUserDetails() {
    var detailsEl = $('userDetails');
    if (!detailsEl) return;
    if (isDemo || !currentUser) { detailsEl.style.display = 'none'; return; }
    detailsEl.style.display = 'flex';

    var expiryValue = $('expiryValue');
    var expirySub = $('expirySub');
    var expiryIcon = $('expiryIcon');
    var expiryIconWrap = $('expiryIconWrap');
    var progressWrap = $('expiryProgressWrap');
    var progressBar = $('expiryProgressBar');
    if (!expiryValue || !expirySub) return;

    if (currentUser.role === 'admin') {
        expiryValue.textContent = 'Vĩnh viễn';
        expiryValue.className = 'detail-value permanent';
        expirySub.textContent = 'Tài khoản quản trị viên';
        expiryIcon.className = 'fas fa-infinity';
        expiryIconWrap.className = 'detail-icon permanent';
        if (progressWrap) progressWrap.style.display = 'none';
        return;
    }

    if (!currentUser.expiresAt) {
        expiryValue.textContent = 'Vĩnh viễn';
        expiryValue.className = 'detail-value permanent';
        expirySub.textContent = 'Không giới hạn thời gian';
        expiryIcon.className = 'fas fa-infinity';
        expiryIconWrap.className = 'detail-icon permanent';
        if (progressWrap) progressWrap.style.display = 'none';
        return;
    }

    var expDate;
    try {
        var ea = currentUser.expiresAt;
        if (typeof ea.toDate === 'function') expDate = ea.toDate();
        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
        else expDate = new Date(ea);
    } catch(e) {
        expiryValue.textContent = '-'; expirySub.textContent = ''; return;
    }
    if (!expDate || isNaN(expDate.getTime())) {
        expiryValue.textContent = '-'; expirySub.textContent = ''; return;
    }

    var now = Date.now();
    var expTime = expDate.getTime();
    var daysLeft = Math.ceil((expTime - now) / (24 * 60 * 60 * 1000));
    var dateStr = expDate.toLocaleDateString('vi-VN');

    expiryValue.className = 'detail-value';
    expiryIconWrap.className = 'detail-icon';

    if (daysLeft < 0) {
        expiryValue.textContent = 'Đã hết hạn';
        expiryValue.classList.add('expired');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Đã hết hạn ' + Math.abs(daysLeft) + ' ngày trước';
        expiryIcon.className = 'fas fa-calendar-times';
        expiryIconWrap.classList.add('expired');
        if (progressWrap) progressWrap.style.display = 'none';
    } else if (daysLeft === 0) {
        expiryValue.textContent = 'Hết hạn hôm nay';
        expiryValue.classList.add('urgent');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Vui lòng liên hệ Admin để gia hạn!';
        expiryIcon.className = 'fas fa-exclamation-circle';
        expiryIconWrap.classList.add('urgent');
        if (progressWrap) progressWrap.style.display = 'none';
    } else if (daysLeft <= 3) {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('urgent');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Sắp hết hạn, vui lòng gia hạn!';
        expiryIcon.className = 'fas fa-exclamation-circle';
        expiryIconWrap.classList.add('urgent');
        var total3 = 30 * 24 * 60 * 60 * 1000;
        var pct3 = Math.max(0, Math.min(100, ((total3 - (expTime - now)) / total3) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar urgent';
            progressBar.style.width = pct3 + '%';
        }
    } else if (daysLeft <= 7) {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('warn');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b>';
        expiryIcon.className = 'fas fa-clock';
        expiryIconWrap.classList.add('warn');
        var total7 = 30 * 24 * 60 * 60 * 1000;
        var pct7 = Math.max(0, Math.min(100, ((total7 - (expTime - now)) / total7) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar warn';
            progressBar.style.width = pct7 + '%';
        }
    } else {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('ok');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b>';
        expiryIcon.className = 'fas fa-calendar-check';
        expiryIconWrap.classList.add('ok');
        var totalOk = 30 * 24 * 60 * 60 * 1000;
        var pctOk = Math.max(0, Math.min(100, ((totalOk - (expTime - now)) / totalOk) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar ok';
            progressBar.style.width = pctOk + '%';
        }
    }
}

/* ============ LOGIN UI ============ */
window.showLoginModal = function() {
    $('loginModal').classList.add('show');
    $('loginError').classList.remove('show');
};
function hideLoginModal() { $('loginModal').classList.remove('show'); }

function showLoginError(msg) {
    var el = $('loginError');
    el.innerHTML = '<i class="fas fa-exclamation-triangle"></i> ' + msg;
    el.classList.add('show');
}

function logLogin(u) {
    try {
        var today = new Date().toDateString();
        var logKey = 'login_log_' + u.email;
        if (localStorage.getItem(logKey) === today) return;
        db.collection('login_logs').add({
            email: u.email, name: u.name, role: u.role,
            time: firebase.firestore.FieldValue.serverTimestamp(),
            userAgent: navigator.userAgent.substring(0, 100)
        }).then(function() {
            try { localStorage.setItem(logKey, today); } catch(e) {}
        }).catch(function() {});
    } catch(e) {}
}

/* ============ ADMIN PANEL ============ */
function initAdminPanel() {
    if ($('openAdminBtn')) {
        $('openAdminBtn').addEventListener('click', function() {
            $('userDropdown').classList.remove('show');
            openAdminPanel();
        });
    }
    if ($('adminClose')) {
        $('adminClose').addEventListener('click', function() {
            $('adminModal').classList.remove('show');
        });
    }
    if ($('adminModal')) {
        $('adminModal').addEventListener('click', function(e) {
            if (e.target === this) $('adminModal').classList.remove('show');
        });
    }
    if ($('refreshUsersBtn')) {
        $('refreshUsersBtn').addEventListener('click', function() {
            try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
            loadUsers(true);
        });
    }
    if ($('exportExcelBtn')) $('exportExcelBtn').addEventListener('click', doExportExcel);
    if ($('importExcelBtn')) $('importExcelBtn').addEventListener('click', function() {
        $('importFileInput').click();
    });
    if ($('importFileInput')) {
        $('importFileInput').addEventListener('change', handleImportFileSelect);
    }
    if ($('importClose')) $('importClose').addEventListener('click', function() {
        $('importModal').classList.remove('show'); importRows = [];
    });
    if ($('importCancelBtn')) $('importCancelBtn').addEventListener('click', function() {
        $('importModal').classList.remove('show'); importRows = [];
    });
    if ($('importModal')) {
        $('importModal').addEventListener('click', function(e) {
            if (e.target === this) { $('importModal').classList.remove('show'); importRows = []; }
        });
    }
    if ($('importConfirmBtn')) $('importConfirmBtn').addEventListener('click', doImport);
    if ($('showAddUserBtn')) {
        $('showAddUserBtn').addEventListener('click', function() {
            $('addUserForm').classList.toggle('show');
            if ($('addUserForm').classList.contains('show')) $('newUserEmail').focus();
        });
    }
    if ($('cancelAddUser')) {
        $('cancelAddUser').addEventListener('click', function() {
            $('addUserForm').classList.remove('show');
            $('newUserEmail').value = '';
            $('newUserName').value = '';
            $('newUserRole').value = 'user';
            $('newUserExpires').value = '';
        });
    }
    if ($('confirmAddUser')) $('confirmAddUser').addEventListener('click', doAddUser);
}

async function doAddUser() {
    var email = $('newUserEmail').value.trim().toLowerCase();
    var name = $('newUserName').value.trim();
    var role = $('newUserRole').value;
    var expiresVal = $('newUserExpires').value;

    if (!email || !email.includes('@')) { alert('Email không hợp lệ'); return; }
    if (!name) name = email.split('@')[0];

    if (role === 'admin' && !isSuperAdmin()) {
        alert('⚠️ Chỉ Super Admin mới có quyền thêm admin!');
        return;
    }

    try {
        var docRef = db.collection('allowed_users').doc(email);
        var doc = await docRef.get();
        if (doc.exists) { alert('Email này đã tồn tại!'); return; }

        var setData = {
            name: name, role: role,
            addedAt: firebase.firestore.FieldValue.serverTimestamp(),
            addedBy: currentUser.email
        };
        if (expiresVal && role !== 'admin') {
            var d = new Date(expiresVal + 'T23:59:59');
            if (!isNaN(d.getTime())) {
                setData.expiresAt = firebase.firestore.Timestamp.fromDate(d);
            }
        }
        await docRef.set(setData);
        $('addUserForm').classList.remove('show');
        $('newUserEmail').value = '';
        $('newUserName').value = '';
        $('newUserRole').value = 'user';
        $('newUserExpires').value = '';
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        loadUsers(true);
    } catch(e) { alert('Lỗi: ' + e.message); }
}

function openAdminPanel() {
    if (!currentUser || currentUser.role !== 'admin') return;
    $('adminModal').classList.add('show');
    loadUsers(false);
    loadLogs();
}

function loadUsers(forceRefresh) {
    var cacheKey = 'admin_users_cache';
    if (forceRefresh) { try { localStorage.removeItem(cacheKey); } catch(e) {} }

    if (!forceRefresh) {
        try {
            var cached = JSON.parse(localStorage.getItem(cacheKey) || 'null');
            if (cached && cached.expires > Date.now() && cached.data && Array.isArray(cached.data)) {
                usersCache = cached.data;
                renderUsers(usersCache);
                renderAdminStats();
                loadLastLoginMap();
                return;
            }
        } catch(e) { try { localStorage.removeItem(cacheKey); } catch(e2) {} }
    }

    $('userList').innerHTML = '<div class="no-data"><i class="fas fa-spinner fa-pulse"></i>Đang tải...</div>';

    db.collection('allowed_users').get()
        .then(function(snapshot) {
            usersCache = [];
            snapshot.forEach(function(doc) {
                var data = doc.data() || {};
                usersCache.push({
                    email: doc.id,
                    name: data.name || '',
                    role: data.role || 'user',
                    expiresAt: data.expiresAt || null
                });
            });
            usersCache.sort(function(a, b) { return (a.email || '').localeCompare(b.email || ''); });
            try {
                localStorage.setItem(cacheKey, JSON.stringify({
                    data: usersCache, expires: Date.now() + 5 * 60 * 1000
                }));
            } catch(e) {}
            renderUsers(usersCache);
            renderAdminStats();
            return loadLastLoginMap();
        })
        .catch(function(err) {
            $('userList').innerHTML = '<div class="no-data" style="color:#dc2626;padding:1rem"><i class="fas fa-exclamation-triangle"></i>Lỗi: ' + err.message + '</div>';
        });
}

function loadLastLoginMap() {
    var hidden = isHiddenAdmin();
    var myEmail = (currentUser && currentUser.email ? currentUser.email.toLowerCase() : '');
    return db.collection('login_logs').orderBy('time', 'desc').limit(500).get()
        .then(function(snapshot) {
            lastLoginMap = {};
            snapshot.forEach(function(doc) {
                var d = doc.data();
                var email = (d.email || '').toLowerCase();
                if (hidden && d.role === 'admin' && email !== myEmail) return;
                if (!lastLoginMap[email] && d.time) lastLoginMap[email] = d.time.toDate();
            });
            renderUsers(usersCache);
        })
        .catch(function() {});
}

function renderAdminStats() {
    var hidden = isHiddenAdmin();
    var usersCacheOnly = usersCache.filter(function(u) { return u.role !== 'admin'; });
    var users = usersCacheOnly.length;
    var admins = usersCache.length - users;
    var now = Date.now();
    var day1 = 24 * 60 * 60 * 1000;
    var day7 = 7 * 24 * 60 * 60 * 1000;
    var active = 0, online = 0, never = 0, expired = 0, expiring = 0;

    usersCacheOnly.forEach(function(u) {
        var last = lastLoginMap[(u.email || '').toLowerCase()];
        if (last) {
            var diff = now - last.getTime();
            if (diff <= day7) active++;
            if (diff <= day1) online++;
        } else { never++; }
        if (u.expiresAt) {
            var d = getExpiryDate(u.expiresAt);
            if (d && !isNaN(d.getTime())) {
                var daysLeft = Math.ceil((d.getTime() - now) / (24 * 60 * 60 * 1000));
                if (daysLeft < 0) expired++;
                else if (daysLeft <= 7) expiring++;
            }
        }
    });

    if (hidden) {
        $('adminStats').innerHTML =
            '<div class="stat-card"><div class="num">' + users + '</div><div class="label">Tổng User</div></div>' +
            '<div class="stat-card"><div class="num" style="color:#16a34a">' + online + '</div><div class="label">Đang hoạt động</div></div>' +
            '<div class="stat-card"><div class="num" style="color:#f59e0b">' + expiring + '</div><div class="label">Sắp hết hạn</div></div>' +
            '<div class="stat-card"><div class="num" style="color:#dc2626">' + expired + '</div><div class="label">Hết hạn</div></div>';
        $('adminUserCount').textContent = users;
        return;
    }

    $('adminStats').innerHTML =
        '<div class="stat-card"><div class="num">' + users + '</div><div class="label">Tổng User</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#16a34a">' + online + '</div><div class="label">Đang hoạt động</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#3b82f6">' + active + '</div><div class="label">Active 7d</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#f59e0b">' + admins + '</div><div class="label">Admin</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#94a3b8">' + never + '</div><div class="label">Chưa login</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#f59e0b">' + expiring + '</div><div class="label">Sắp hết hạn</div></div>' +
        '<div class="stat-card"><div class="num" style="color:#dc2626">' + expired + '</div><div class="label">Hết hạn</div></div>';
    $('adminUserCount').textContent = users;
}

function renderUsers(items) {
    var list = $('userList');
    var hidden = isHiddenAdmin();
    var superAdmin = isSuperAdmin();
    var myEmail = (currentUser && currentUser.email ? currentUser.email.toLowerCase() : '');

    var displayItems = items;
    if (hidden) {
        displayItems = items.filter(function(u) {
            var uEmail = (u.email || '').toLowerCase();
            if (u.role === 'admin' && uEmail !== myEmail) return false;
            return true;
        });
    }

    if (!displayItems.length) {
        list.innerHTML = '<div class="no-data" style="padding:1.5rem;font-size:.85rem"><i class="fas fa-search"></i>Không có user nào</div>';
        return;
    }

    var now = Date.now();
    var day7 = 7 * 24 * 60 * 60 * 1000;
    var day30 = 30 * 24 * 60 * 60 * 1000;

    list.innerHTML = displayItems.map(function(u) {
        var isMe = u.email === currentUser.email;
        var isAdmin = u.role === 'admin';
        var targetIsSuper = (u.email || '').toLowerCase() === SUPER_ADMIN.toLowerCase();
        var canModifyAdmin = superAdmin && isAdmin && !isMe && !targetIsSuper;

        var roleBtn = '';
        if (isAdmin) {
            if (canModifyAdmin) {
                roleBtn = '<button class="u-btn" onclick="changeRole(\'' + escapeJs(u.email) + '\', \'user\')" title="Hạ xuống User"><i class="fas fa-user"></i></button>';
            } else {
                var reason = isMe ? 'Không thể tự hạ quyền chính mình' : (targetIsSuper ? 'Không thể hạ quyền Super Admin' : 'Chỉ Super Admin mới hạ quyền được');
                roleBtn = '<button class="u-btn" disabled title="' + escapeHtml(reason) + '"><i class="fas fa-user"></i></button>';
            }
        } else {
            if (superAdmin) {
                roleBtn = '<button class="u-btn" onclick="changeRole(\'' + escapeJs(u.email) + '\', \'admin\')" title="Nâng lên Admin"><i class="fas fa-shield-alt"></i></button>';
            } else {
                roleBtn = '<button class="u-btn" disabled title="Chỉ Super Admin mới nâng quyền được"><i class="fas fa-shield-alt"></i></button>';
            }
        }

        var deleteBtn = '';
        if (isMe) {
            deleteBtn = '<button class="u-btn danger" disabled title="Không thể tự xóa chính mình"><i class="fas fa-trash"></i></button>';
        } else if (targetIsSuper) {
            deleteBtn = '<button class="u-btn danger" disabled title="Không thể xóa Super Admin"><i class="fas fa-trash"></i></button>';
        } else if (isAdmin) {
            if (canModifyAdmin) {
                deleteBtn = '<button class="u-btn danger" onclick="deleteUser(\'' + escapeJs(u.email) + '\')" title="Xóa admin"><i class="fas fa-trash"></i></button>';
            } else {
                deleteBtn = '<button class="u-btn danger" disabled title="Chỉ Super Admin mới xóa được admin"><i class="fas fa-trash"></i></button>';
            }
        } else {
            deleteBtn = '<button class="u-btn danger" onclick="deleteUser(\'' + escapeJs(u.email) + '\')" title="Xóa"><i class="fas fa-trash"></i></button>';
        }

        var expiryBtn = '';
        if (!isAdmin) {
            var btnCls = 'u-btn expiry';
            var tooltip = 'Chỉnh hạn sử dụng';
            if (u.expiresAt) {
                var d = getExpiryDate(u.expiresAt);
                if (d && !isNaN(d.getTime())) {
                    var daysLeftExp = Math.ceil((d.getTime() - Date.now()) / (24 * 60 * 60 * 1000));
                    tooltip = 'Chỉnh hạn (hiện tại: ' + d.toLocaleDateString('vi-VN') + ', còn ' + Math.max(0, daysLeftExp) + ' ngày)';
                    if (daysLeftExp <= 7) btnCls += ' urgent';
                }
            } else {
                tooltip = 'Chỉnh hạn (hiện tại: Vĩnh viễn)';
            }
            expiryBtn = '<button class="' + btnCls + '" onclick="openEditExpiry(\'' + escapeJs(u.email) + '\')" title="' + escapeHtml(tooltip) + '"><i class="fas fa-calendar-alt"></i></button>';
        }

        var lastLoginHtml = '';
        var last = lastLoginMap[(u.email || '').toLowerCase()];
        if (last) {
            var diff = now - last.getTime();
            var cls = diff <= day7 ? 'active' : (diff <= day30 ? 'recent' : '');
            lastLoginHtml = '<div class="u-last-login ' + cls + '"><i class="fas fa-clock"></i> ' + formatTimeDiff(diff) + '</div>';
        } else {
            lastLoginHtml = '<div class="u-last-login"><i class="fas fa-times-circle"></i> Chưa đăng nhập</div>';
        }

        var expiryHtml = '';
        if (isAdmin) {
            expiryHtml = '<div class="u-expiry permanent"><i class="fas fa-infinity"></i> Vĩnh viễn</div>';
        } else if (!u.expiresAt) {
            expiryHtml = '<div class="u-expiry permanent" onclick="openEditExpiry(\'' + escapeJs(u.email) + '\')" title="Click để chỉnh"><i class="fas fa-infinity"></i> Vĩnh viễn</div>';
        } else {
            var expDate = getExpiryDate(u.expiresAt);
            if (expDate && !isNaN(expDate.getTime())) {
                var daysLeft = Math.ceil((expDate.getTime() - now) / (24 * 60 * 60 * 1000));
                var expCls = 'ok', expIcon = 'fa-calendar-check', expText = '';
                if (daysLeft < 0) { expCls = 'expired'; expIcon = 'fa-calendar-times'; expText = 'Hết hạn ' + Math.abs(daysLeft) + ' ngày trước'; }
                else if (daysLeft === 0) { expCls = 'urgent'; expIcon = 'fa-exclamation-circle'; expText = 'Hết hạn hôm nay'; }
                else if (daysLeft <= 3) { expCls = 'urgent'; expIcon = 'fa-exclamation-circle'; expText = 'Còn ' + daysLeft + ' ngày'; }
                else if (daysLeft <= 7) { expCls = 'warn'; expIcon = 'fa-clock'; expText = 'Còn ' + daysLeft + ' ngày'; }
                else { expText = 'Còn ' + daysLeft + ' ngày'; }
                expiryHtml = '<div class="u-expiry ' + expCls + '" onclick="openEditExpiry(\'' + escapeJs(u.email) + '\')" title="Click để chỉnh"><i class="fas ' + expIcon + '"></i> ' + escapeHtml(expText) + ' • ' + expDate.toLocaleDateString('vi-VN') + '</div>';
            }
        }

        var roleBadge = isAdmin
            ? '<span class="u-role ' + (targetIsSuper ? 'super' : 'admin') + '">' + (targetIsSuper ? '👑 super' : 'admin') + '</span>'
            : '<span class="u-role user">user</span>';

        return '<div class="user-row" data-email="' + escapeHtml(u.email) + '">' +
            '<div class="u-info">' +
                '<div class="u-name">' + escapeHtml(u.name || u.email.split('@')[0]) + (isMe ? ' <span style="color:#94a3b8;font-size:.7rem">(bạn)</span>' : '') + '</div>' +
                '<div class="u-email">' + escapeHtml(u.email) + '</div>' +
                lastLoginHtml + expiryHtml +
            '</div>' + roleBadge +
            '<div class="u-actions">' + expiryBtn + roleBtn + deleteBtn + '</div>' +
        '</div>';
    }).join('');
}

window.changeRole = async function(email, newRole) {
    var target = usersCache.find(function(u) { return u.email === email; });
    if (!target) { alert('Không tìm thấy user!'); return; }
    var isMe = email === currentUser.email;
    var isAdmin = target.role === 'admin';
    var superAdmin = isSuperAdmin();
    var targetIsSuper = (email || '').toLowerCase() === SUPER_ADMIN.toLowerCase();
    if (isMe && newRole === 'user') { alert('⚠️ Không thể tự hạ quyền admin của chính mình!'); return; }
    if (targetIsSuper) { alert('⚠️ Không thể thay đổi quyền của Super Admin!'); return; }
    if (isAdmin && newRole === 'user' && !superAdmin) { alert('⚠️ Chỉ Super Admin mới có quyền hạ cấp admin khác!'); return; }
    if (!isAdmin && newRole === 'admin' && !superAdmin) { alert('⚠️ Chỉ Super Admin mới có quyền nâng cấp lên admin!'); return; }
    var action = newRole === 'admin' ? 'NÂNG LÊN ADMIN' : 'HẠ XUỐNG USER';
    if (!confirm(action + ' cho tài khoản:\n\n' + email + '\n\nBạn có chắc không?')) return;
    try {
        await db.collection('allowed_users').doc(email).update({ role: newRole });
        try { localStorage.removeItem('user_cache_' + email); } catch(e) {}
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        loadUsers(true);
    } catch(e) { alert('Lỗi: ' + e.message); }
};

window.deleteUser = async function(email) {
    var target = usersCache.find(function(u) { return u.email === email; });
    if (!target) { alert('Không tìm thấy user!'); return; }
    var isMe = email === currentUser.email;
    var isAdmin = target.role === 'admin';
    var superAdmin = isSuperAdmin();
    var targetIsSuper = (email || '').toLowerCase() === SUPER_ADMIN.toLowerCase();
    if (isMe) { alert('⚠️ Không thể tự xóa tài khoản của chính mình!'); return; }
    if (targetIsSuper) { alert('⚠️ Không thể xóa Super Admin!'); return; }
    if (isAdmin && !superAdmin) { alert('⚠️ Chỉ Super Admin mới có quyền xóa admin khác!'); return; }
    var confirmMsg = isAdmin
        ? '⚠️ XÓA ADMIN\n\n' + email + '\n\nNgười này sẽ mất quyền quản trị và không đăng nhập được nữa.\n\nBạn có chắc không?'
        : '⚠️ XÓA TÀI KHOẢN\n\n' + email + '\n\nNgười này sẽ không đăng nhập được nữa.\n\nBạn có chắc không?';
    if (!confirm(confirmMsg)) return;
    try {
        await db.collection('allowed_users').doc(email).delete();
        try { localStorage.removeItem('user_cache_' + email); } catch(e) {}
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        loadUsers(true);
    } catch(e) { alert('Lỗi: ' + e.message); }
};

function loadLogs() {
    var hidden = isHiddenAdmin();
    var logTitleEl = $('logsTitle');
    if (hidden) {
        if (logTitleEl) logTitleEl.style.display = 'none';
        $('logsList').style.display = 'none';
        $('logsList').innerHTML = '';
        return;
    }
    if (logTitleEl) logTitleEl.style.display = 'flex';
    $('logsList').style.display = 'block';
    db.collection('login_logs').orderBy('time', 'desc').limit(30).get()
        .then(function(snapshot) {
            if (snapshot.empty) {
                $('logsList').innerHTML = '<div class="no-data" style="padding:1rem;font-size:.8rem">Chưa có log</div>';
                return;
            }
            var html = '';
            snapshot.forEach(function(doc) {
                var d = doc.data();
                var time = d.time ? new Date(d.time.toDate()).toLocaleString('vi-VN') : 'N/A';
                html += '<div class="log-item">' +
                    '<span class="log-time">' + time + '</span>' +
                    '<span class="log-msg"><b>' + escapeHtml(d.name || d.email) + '</b> (' + escapeHtml(d.role || 'user') + ')</span>' +
                '</div>';
            });
            $('logsList').innerHTML = html;
        })
        .catch(function(err) {
            $('logsList').innerHTML = '<div class="no-data" style="padding:1rem;font-size:.8rem;color:#dc2626">Lỗi: ' + err.message + '</div>';
        });
}

/* ============ EXPIRY EDIT ============ */
window.openEditExpiry = function(email) {
    var user = usersCache.find(function(u) { return u.email === email; });
    if (!user) { alert('Không tìm thấy user!'); return; }
    if (user.role === 'admin') { alert('Admin có hạn vĩnh viễn, không cần chỉnh!'); return; }
    editingExpiryEmail = email;
    $('editExpiryName').textContent = user.name || email.split('@')[0];
    $('editExpiryEmail').textContent = email;
    if (user.expiresAt) {
        var d = getExpiryDate(user.expiresAt);
        if (d && !isNaN(d.getTime())) $('editExpiryInput').value = formatDate(d);
        else $('editExpiryInput').value = '';
    } else {
        $('editExpiryInput').value = '';
    }
    $('editExpiryModal').classList.add('show');
};

window.setQuickExpiry = function(days) {
    var d = new Date();
    d.setDate(d.getDate() + days);
    d.setHours(23, 59, 59);
    $('editExpiryInput').value = formatDate(d);
};
window.setQuickExpiryPermanent = function() { $('editExpiryInput').value = ''; };

async function doUpdateExpiry() {
    if (!editingExpiryEmail) return;
    var dateVal = $('editExpiryInput').value;
    var updateData = {};
    if (dateVal) {
        var d = new Date(dateVal + 'T23:59:59');
        if (isNaN(d.getTime())) { alert('Ngày không hợp lệ!'); return; }
        updateData.expiresAt = firebase.firestore.Timestamp.fromDate(d);
    } else {
        updateData.expiresAt = null;
    }
    var btn = $('editExpiryConfirm');
    btn.disabled = true;
    var originalHtml = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Đang lưu...';
    try {
        await db.collection('allowed_users').doc(editingExpiryEmail).update(updateData);
        try { localStorage.removeItem('user_cache_' + editingExpiryEmail); } catch(e) {}
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        $('editExpiryModal').classList.remove('show');
        var msg = dateVal
            ? '✅ Đã cập nhật hạn đến ngày ' + new Date(dateVal).toLocaleDateString('vi-VN')
            : '✅ Đã đặt thành vĩnh viễn';
        alert(msg);
        loadUsers(true);
    } catch(err) {
        alert('❌ Lỗi: ' + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalHtml;
    }
}

/* ============ CHANGE NAME ============ */
async function doChangeName() {
    if (!editingEmail) return;
    var newName = $('changeNameInput').value.trim();
    if (!newName) { alert('Tên không được để trống!'); return; }
    if (newName.length > 50) { alert('Tên quá dài!'); return; }
    var btn = $('changeNameConfirm');
    btn.disabled = true;
    var originalHtml = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Đang lưu...';
    try {
        await db.collection('allowed_users').doc(editingEmail).update({ name: newName });
        try { localStorage.removeItem('user_cache_' + editingEmail); } catch(e) {}
        try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
        if (currentUser && currentUser.email === editingEmail) {
            currentUser.name = newName;
            $('userName').textContent = newName;
            var avatar = $('userAvatar');
            if (!currentUser.photo) {
                avatar.src = 'data:image/svg+xml;utf8,' + encodeURIComponent(
                    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">' +
                    '<rect fill="#2563eb" width="100" height="100"/>' +
                    '<text x="50" y="65" font-size="45" fill="#fff" text-anchor="middle" font-family="sans-serif" font-weight="bold">' +
                    newName.charAt(0).toUpperCase() + '</text></svg>'
                );
            }
            try {
                localStorage.setItem('user_cache_' + editingEmail, JSON.stringify({
                    data: currentUser, expires: Date.now() + 12 * 60 * 60 * 1000
                }));
            } catch(e) {}
        }
        $('changeNameModal').classList.remove('show');
        alert('✅ Đã đổi tên thành công!');
        if ($('adminModal').classList.contains('show')) loadUsers(true);
    } catch(err) {
        alert('❌ Lỗi: ' + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalHtml;
    }
}

/* ============ EXPORT EXCEL ============ */
function doExportExcel() {
    var usersOnly = usersCache.filter(function(u) { return u.role !== 'admin'; });
    if (!usersOnly.length) {
        alert('Không có user nào để export!\n(Admin không được export)');
        return;
    }
    try {
        var wb = XLSX.utils.book_new();
        var now = Date.now();
        usersOnly.sort(function(a, b) {
            var da = getExpiryTimestamp(a.expiresAt);
            var db = getExpiryTimestamp(b.expiresAt);
            if (da === null && db === null) return 0;
            if (da === null) return 1;
            if (db === null) return -1;
            return da - db;
        });
        var COLUMNS = [
            { header: 'STT', width: 6 }, { header: 'email', width: 35 },
            { header: 'name', width: 25 }, { header: 'phone', width: 16 },
            { header: 'expiresAt', width: 14 }, { header: 'Trạng thái', width: 22 },
            { header: 'Ghi chú', width: 25 },
        ];
        var aoa = [COLUMNS.map(function(c) { return c.header; })];
        var stats = { total: usersOnly.length, permanent: 0, expired: 0, urgent: 0, warning: 0, ok: 0 };
        var statusTypes = [];

        usersOnly.forEach(function(u) {
            var expDate = getExpiryDate(u.expiresAt);
            var expStr = '', statusStr = '', statusType = 'ok';
            if (!expDate) {
                statusStr = '∞ Vĩnh viễn'; statusType = 'permanent'; stats.permanent++;
            } else {
                expStr = formatDate(expDate);
                var daysLeft = Math.ceil((expDate.getTime() - now) / (24 * 60 * 60 * 1000));
                if (daysLeft < 0) { statusStr = '❌ Hết hạn ' + Math.abs(daysLeft) + ' ngày'; statusType = 'expired'; stats.expired++; }
                else if (daysLeft === 0) { statusStr = '⏰ Hết hạn hôm nay'; statusType = 'urgent'; stats.urgent++; }
                else if (daysLeft <= 3) { statusStr = '🔴 Còn ' + daysLeft + ' ngày'; statusType = 'urgent'; stats.urgent++; }
                else if (daysLeft <= 7) { statusStr = '🟡 Còn ' + daysLeft + ' ngày'; statusType = 'warning'; stats.warning++; }
                else { statusStr = '🟢 Còn ' + daysLeft + ' ngày'; statusType = 'ok'; stats.ok++; }
            }
            statusTypes.push(statusType);
            aoa.push(['', u.email || '', u.name || '', '', expStr, statusStr, '']);
        });

        var totalRows = aoa.length;
        var ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols'] = COLUMNS.map(function(c) { return { wch: c.width }; });
        ws['!rows'] = [{ hpt: 30 }];
        for (var r = 1; r < totalRows; r++) ws['!rows'].push({ hpt: 22 });
        ws['!freeze'] = { xSplit: 0, ySplit: 1 };
        ws['!autofilter'] = { ref: 'A1:G' + totalRows };

        var headerStyle = {
            font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
            fill: { fgColor: { rgb: '2563EB' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
                top: { style: 'thin', color: { rgb: '1E40AF' } }, bottom: { style: 'thin', color: { rgb: '1E40AF' } },
                left: { style: 'thin', color: { rgb: '1E40AF' } }, right: { style: 'thin', color: { rgb: '1E40AF' } }
            }
        };
        ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1'].forEach(function(ref) {
            if (ws[ref]) ws[ref].s = headerStyle;
        });

        function getBorderStyle() {
            return {
                top: { style: 'thin', color: { rgb: 'CBD5E1' } }, bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
                left: { style: 'thin', color: { rgb: 'CBD5E1' } }, right: { style: 'thin', color: { rgb: 'CBD5E1' } }
            };
        }

        for (var r = 2; r <= totalRows; r++) {
            var statusType = statusTypes[r - 2] || 'ok';
            var sttRef = 'A' + r;
            if (!ws[sttRef]) ws[sttRef] = { v: '', t: 's' };
            ws[sttRef].f = 'IF(B' + r + '<>"",ROW()-1,"")';
            ws[sttRef].t = 'n';
            ws[sttRef].s = { font: { bold: true, color: { rgb: '64748B' }, sz: 10 }, fill: { fgColor: { rgb: 'F1F5F9' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: getBorderStyle() };
            var emailRef = 'B' + r; if (!ws[emailRef]) ws[emailRef] = { v: '', t: 's' };
            ws[emailRef].s = { fill: { fgColor: { rgb: 'DBEAFE' } }, alignment: { horizontal: 'left', vertical: 'center' }, border: getBorderStyle() };
            var nameRef = 'C' + r; if (!ws[nameRef]) ws[nameRef] = { v: '', t: 's' };
            ws[nameRef].s = { fill: { fgColor: { rgb: 'F0F9FF' } }, alignment: { horizontal: 'left', vertical: 'center' }, border: getBorderStyle() };
            var phoneRef = 'D' + r; if (!ws[phoneRef]) ws[phoneRef] = { v: '', t: 's' };
            ws[phoneRef].s = { fill: { fgColor: { rgb: 'FEF3C7' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: getBorderStyle() };
            ws[phoneRef].t = 's'; ws[phoneRef].z = '@';
            var expRef = 'E' + r; if (!ws[expRef]) ws[expRef] = { v: '', t: 's' };
            var expBgColor = 'DCFCE7';
            if (statusType === 'expired') expBgColor = 'FEE2E2';
            else if (statusType === 'urgent') expBgColor = 'FECACA';
            else if (statusType === 'warning') expBgColor = 'FEF3C7';
            else if (statusType === 'permanent') expBgColor = 'F1F5F9';
            ws[expRef].s = { fill: { fgColor: { rgb: expBgColor } }, alignment: { horizontal: 'center', vertical: 'center' }, border: getBorderStyle(), font: { bold: statusType === 'expired' || statusType === 'urgent', color: { rgb: statusType === 'expired' ? 'DC2626' : '0F172A' }, sz: 10 } };
            var sttStatusRef = 'F' + r; if (!ws[sttStatusRef]) ws[sttStatusRef] = { v: '', t: 's' };
            var statusBgColor = 'DCFCE7', statusFontColor = '16A34A';
            if (statusType === 'expired') { statusBgColor = 'FEE2E2'; statusFontColor = 'DC2626'; }
            else if (statusType === 'urgent') { statusBgColor = 'FECACA'; statusFontColor = 'DC2626'; }
            else if (statusType === 'warning') { statusBgColor = 'FEF3C7'; statusFontColor = '92400E'; }
            else if (statusType === 'permanent') { statusBgColor = 'DBEAFE'; statusFontColor = '1D4ED8'; }
            ws[sttStatusRef].s = { fill: { fgColor: { rgb: statusBgColor } }, alignment: { horizontal: 'center', vertical: 'center' }, border: getBorderStyle(), font: { bold: true, color: { rgb: statusFontColor }, sz: 10 } };
            var noteRef = 'G' + r; if (!ws[noteRef]) ws[noteRef] = { v: '', t: 's' };
            ws[noteRef].s = { fill: { fgColor: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'left', vertical: 'center' }, border: getBorderStyle() };
        }

        XLSX.utils.book_append_sheet(wb, ws, 'Users');

        var ws2 = XLSX.utils.aoa_to_sheet([
            ['📊  THỐNG KÊ TÀI KHOẢN', '', ''],
            ['', '', ''],
            ['Tổng số user', stats.total, ''],
            ['', '', ''],
            ['🟢 Còn nhiều thời gian (> 7 ngày)', stats.ok, ''],
            ['🟡 Sắp hết hạn (4-7 ngày)', stats.warning, ''],
            ['🔴 Sắp hết hạn (≤ 3 ngày)', stats.urgent, ''],
            ['❌ Đã hết hạn', stats.expired, ''],
            ['∞  Vĩnh viễn', stats.permanent, ''],
            ['', '', ''],
            ['📅 Ngày export', new Date().toLocaleString('vi-VN'), ''],
        ]);
        ws2['!cols'] = [{ wch: 35 }, { wch: 18 }, { wch: 20 }];
        if (ws2['A1']) ws2['A1'].s = { font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '2563EB' } }, alignment: { horizontal: 'center', vertical: 'center' } };
        [3, 5, 6, 7, 8, 9].forEach(function(r) {
            var aRef = 'A' + r, bRef = 'B' + r;
            if (ws2[aRef]) ws2[aRef].s = { font: { bold: true, sz: 11, color: { rgb: '0F172A' } }, alignment: { horizontal: 'left', vertical: 'center', indent: 1 } };
            if (ws2[bRef]) ws2[bRef].s = { font: { bold: true, sz: 14 }, alignment: { horizontal: 'center', vertical: 'center' } };
        });
        if (ws2['B5']) ws2['B5'].s.font = { bold: true, sz: 14, color: { rgb: '16A34A' } };
        if (ws2['B6']) ws2['B6'].s.font = { bold: true, sz: 14, color: { rgb: 'D97706' } };
        if (ws2['B7']) ws2['B7'].s.font = { bold: true, sz: 14, color: { rgb: 'DC2626' } };
        if (ws2['B8']) ws2['B8'].s.font = { bold: true, sz: 14, color: { rgb: 'DC2626' } };
        if (ws2['B9']) ws2['B9'].s.font = { bold: true, sz: 14, color: { rgb: '2563EB' } };
        XLSX.utils.book_append_sheet(wb, ws2, 'Thống kê');

        var ws3 = XLSX.utils.aoa_to_sheet([
            ['📖  HƯỚNG DẪN SỬ DỤNG FILE EXPORT', '', ''],
            ['', '', ''],
            ['🎯  MỤC ĐÍCH', '', ''],
            ['', 'File này là bản backup danh sách user', ''],
            ['', 'Có thể import lại để khôi phục', ''],
            ['', '', ''],
            ['📌  CỘT QUAN TRỌNG', '', ''],
            ['', 'email', '✅ BẮT BUỘC khi import lại'],
            ['', 'name', '⭕ Tùy chọn'],
            ['', 'expiresAt', '⭕ Định dạng YYYY-MM-DD. Trống = vĩnh viễn'],
            ['', '', ''],
            ['🗑️  CỘT CHỈ ĐỂ THAM KHẢO', '', ''],
            ['', 'STT', 'Tự động đánh số'],
            ['', 'phone', 'Không import'],
            ['', 'Trạng thái', 'Tự động tính theo ngày hiện tại'],
            ['', 'Ghi chú', 'Không import'],
            ['', '', ''],
            ['⚠️  LƯU Ý', '', ''],
            ['', '🚫 Admin không có trong file', 'Chỉ export user thường'],
            ['', '📅 Định dạng ngày', 'YYYY-MM-DD'],
            ['', '🔄 Sắp xếp', 'Sắp hết hạn lên đầu, vĩnh viễn xuống cuối'],
            ['', '🎨 Màu sắc', 'Đỏ = hết hạn, Cam = gấp, Vàng = sắp hết, Xanh = ổn'],
            ['', '', ''],
            ['🚀  CÁCH IMPORT LẠI', '', ''],
            ['', '1.', 'Mở web, đăng nhập Admin'],
            ['', '2.', 'Click avatar → Quản lý tài khoản'],
            ['', '3.', 'Nhấn nút "Import"'],
            ['', '4.', 'Chọn file này'],
            ['', '5.', 'Kiểm tra preview → Confirm'],
        ]);
        ws3['!cols'] = [{ wch: 4 }, { wch: 30 }, { wch: 60 }];
        ws3['!rows'] = [{ hpt: 42 }, { hpt: 8 }];
        if (ws3['A1']) ws3['A1'].s = { font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '2563EB' } }, alignment: { horizontal: 'center', vertical: 'center' } };
        ['A3', 'A7', 'A12', 'A18', 'A24'].forEach(function(ref) {
            if (ws3[ref]) ws3[ref].s = { font: { bold: true, sz: 12, color: { rgb: '1E40AF' } }, fill: { fgColor: { rgb: 'DBEAFE' } }, alignment: { horizontal: 'left', vertical: 'center', indent: 1 } };
        });
        for (var rr = 4; rr <= 30; rr++) {
            var bRef = 'B' + rr, cRef = 'C' + rr;
            if (ws3[bRef] && ws3[bRef].v) ws3[bRef].s = { font: { bold: true, sz: 10, color: { rgb: '0F172A' } }, alignment: { horizontal: 'left', vertical: 'center' } };
            if (ws3[cRef] && ws3[cRef].v) ws3[cRef].s = { font: { sz: 10, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } };
        }
        XLSX.utils.book_append_sheet(wb, ws3, 'Hướng dẫn');

        var today = new Date();
        var dateStr = today.getFullYear() + String(today.getMonth() + 1).padStart(2, '0') + String(today.getDate()).padStart(2, '0') + '_' + String(today.getHours()).padStart(2, '0') + String(today.getMinutes()).padStart(2, '0');
        XLSX.writeFile(wb, 'users_export_' + dateStr + '.xlsx', { bookType: 'xlsx', cellStyles: true });
    } catch(err) {
        console.error('Lỗi export:', err);
        alert('❌ Lỗi export: ' + err.message);
    }
}

/* ============ IMPORT EXCEL ============ */
function handleImportFileSelect(e) {
    var file = e.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function(evt) {
        try {
            var data = new Uint8Array(evt.target.result);
            var workbook = XLSX.read(data, { type: 'array' });
            var sheet = workbook.Sheets[workbook.SheetNames[0]];
            var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            processImport(rows);
        } catch(err) {
            alert('❌ Lỗi đọc file: ' + err.message);
        }
        e.target.value = '';
    };
    reader.readAsArrayBuffer(file);
}

function processImport(rows) {
    if (!rows || rows.length < 2) { alert('❌ File rỗng hoặc thiếu header!'); return; }
    var headerRowIdx = -1;
    for (var i = 0; i < Math.min(5, rows.length); i++) {
        var r = rows[i].map(function(c) { return String(c || '').toLowerCase().trim(); });
        if (r.indexOf('email') !== -1) { headerRowIdx = i; break; }
    }
    if (headerRowIdx === -1) {
        alert('❌ Không tìm thấy cột "email"!\n\nFile Excel cần có ít nhất cột "email".');
        return;
    }
    var header = rows[headerRowIdx].map(function(c) { return String(c || '').toLowerCase().trim(); });
    var emailCol = header.indexOf('email');
    var nameCol = header.indexOf('name');
    var expCol = header.indexOf('expiresat');

    var existingUserMap = {};
    var existingAdminSet = {};
    usersCache.forEach(function(u) {
        var email = (u.email || '').toLowerCase();
        if (u.role === 'admin') existingAdminSet[email] = true;
        else existingUserMap[email] = u;
    });

    importRows = [];
    var stats = { total: 0, newUser: 0, update: 0, invalid: 0, skippedAdmin: 0 };
    var seenInFile = {};

    for (var i = headerRowIdx + 1; i < rows.length; i++) {
        var row = rows[i];
        if (!row || row.length === 0) continue;
        var email = emailCol >= 0 ? String(row[emailCol] || '').trim().toLowerCase() : '';
        var name = nameCol >= 0 ? String(row[nameCol] || '').trim() : '';
        var expRaw = expCol >= 0 ? row[expCol] : '';
        if (!email && !name) continue;
        if (email.indexOf('←') === 0 || email.indexOf('•') === 0 || email.indexOf('xóa dòng') !== -1 || email.indexOf('#') === 0 || email.indexOf('⚠') === 0 || email.indexOf('ví dụ') === 0) continue;
        if (existingAdminSet[email]) { stats.skippedAdmin++; continue; }

        var status = 'ok';
        var reason = '';
        var isUpdate = !!existingUserMap[email];
        if (!email) { status = 'error'; reason = 'Thiếu email'; stats.invalid++; }
        else if (!email.includes('@') || !email.includes('.')) { status = 'error'; reason = 'Email không hợp lệ'; stats.invalid++; }
        else if (seenInFile[email]) { status = 'error'; reason = 'Trùng trong file'; stats.invalid++; }
        else if (isUpdate) { stats.update++; seenInFile[email] = true; }
        else { stats.newUser++; seenInFile[email] = true; }

        if (!name && email.indexOf('@') > 0) name = email.split('@')[0];

        var expDate = null;
        var expStr = '';
        if (expRaw) {
            var raw = expRaw;
            if (typeof raw === 'number' && raw > 25569) {
                var d = new Date((raw - 25569) * 86400 * 1000);
                if (!isNaN(d.getTime())) {
                    expDate = d;
                    expStr = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
                }
            } else {
                var s = String(raw).trim();
                if (s) {
                    var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
                    if (m) {
                        var d2 = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]), 23, 59, 59);
                        if (!isNaN(d2.getTime())) {
                            expDate = d2;
                            expStr = m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');
                        }
                    } else {
                        var d3 = new Date(s);
                        if (!isNaN(d3.getTime())) {
                            expDate = d3;
                            expStr = d3.getFullYear() + '-' + String(d3.getMonth() + 1).padStart(2, '0') + '-' + String(d3.getDate()).padStart(2, '0');
                        } else {
                            if (status === 'ok') { status = 'warn'; reason = 'Ngày không hợp lệ (bỏ qua hạn)'; }
                            expDate = null;
                        }
                    }
                }
            }
        }

        importRows.push({
            rowNum: i + 1, email: email, name: name, role: 'user',
            expDate: expDate, expStr: expStr, status: status, reason: reason, isUpdate: isUpdate
        });
    }

    if (importRows.length === 0) {
        var msg = '❌ Không có dòng dữ liệu hợp lệ!';
        if (stats.skippedAdmin > 0) msg += '\n\n(Bỏ qua ' + stats.skippedAdmin + ' admin)';
        alert(msg);
        return;
    }
    renderImportPreview();
    $('importModal').classList.add('show');
}

function renderImportPreview() {
    var tbody = $('importTableBody');
    if (!tbody) return;
    var html = '';
    var countOk = 0, countUpdate = 0, countWarn = 0, countErr = 0;
    importRows.forEach(function(r) {
        var rowCls = '';
        var statusHtml = '';
        if (r.status === 'ok' && r.isUpdate) { rowCls = 'row-update'; statusHtml = '<span class="status-badge update"><i class="fas fa-sync-alt"></i> Cập nhật</span>'; countUpdate++; }
        else if (r.status === 'ok') { rowCls = 'row-new'; statusHtml = '<span class="status-badge ok"><i class="fas fa-plus"></i> Thêm mới</span>'; countOk++; }
        else if (r.status === 'warn') { rowCls = 'row-warn'; statusHtml = '<span class="status-badge warn"><i class="fas fa-exclamation-triangle"></i> ' + escapeHtml(r.reason) + '</span>'; countWarn++; }
        else { rowCls = 'row-error'; statusHtml = '<span class="status-badge err"><i class="fas fa-times"></i> ' + escapeHtml(r.reason) + '</span>'; countErr++; }
        var expDisplay = '—';
        if (r.expStr) {
            var d = new Date(r.expStr + 'T23:59:59');
            var daysLeft = Math.ceil((d.getTime() - Date.now()) / (24 * 60 * 60 * 1000));
            var expColor = daysLeft < 0 ? '#dc2626' : (daysLeft <= 7 ? '#f59e0b' : '#16a34a');
            expDisplay = '<span style="color:' + expColor + ';font-size:.7rem;font-weight:600;">' + r.expStr + '</span>';
        } else {
            expDisplay = '<span style="color:#94a3b8;font-size:.7rem;">Vĩnh viễn</span>';
        }
        html += '<tr class="' + rowCls + '">' +
            '<td>' + r.rowNum + '</td>' +
            '<td><b>' + escapeHtml(r.email) + '</b></td>' +
            '<td>' + escapeHtml(r.name) + '</td>' +
            '<td><span class="role-badge user">user</span></td>' +
            '<td>' + expDisplay + '</td>' +
            '<td>' + statusHtml + '</td>' +
        '</tr>';
    });
    tbody.innerHTML = html;

    var summaryEl = $('importSummary');
    if (summaryEl) {
        summaryEl.innerHTML =
            '<div class="import-stat"><div class="num">' + importRows.length + '</div><div class="label">Tổng</div></div>' +
            '<div class="import-stat ok"><div class="num">' + countOk + '</div><div class="label">Thêm mới</div></div>' +
            '<div class="import-stat update"><div class="num">' + countUpdate + '</div><div class="label">Cập nhật</div></div>' +
            '<div class="import-stat warn"><div class="num">' + countWarn + '</div><div class="label">Cảnh báo</div></div>' +
            '<div class="import-stat err"><div class="num">' + countErr + '</div><div class="label">Lỗi</div></div>';
    }
    var totalImportable = countOk + countUpdate + countWarn;
    if ($('importCount')) $('importCount').textContent = totalImportable;
    if ($('importConfirmBtn')) $('importConfirmBtn').disabled = totalImportable === 0;
}

async function doImport() {
    var skipDuplicates = $('importSkipDuplicates').checked;
    var skipInvalid = $('importSkipInvalid').checked;
    var toImport = importRows.filter(function(r) {
        if (r.status === 'error') return false;
        if (r.status === 'warn' && skipInvalid) return false;
        if (r.isUpdate && skipDuplicates) return false;
        return true;
    });
    if (toImport.length === 0) { alert('⚠️ Không có user nào để import!'); return; }
    var totalNew = toImport.filter(function(r) { return !r.isUpdate; }).length;
    var totalUpdate = toImport.filter(function(r) { return r.isUpdate; }).length;
    if (!confirm('📥 IMPORT ' + toImport.length + ' TÀI KHOẢN?\n\n• Thêm mới: ' + totalNew + '\n• Cập nhật: ' + totalUpdate + '\n\n(Chỉ import user, KHÔNG import admin)\n\nBạn có chắc không?')) return;

    var btn = $('importConfirmBtn');
    if (!btn) return;
    var originalHTML = btn.innerHTML;
    var originalDisabled = btn.disabled;
    btn.disabled = true;
    var icon = btn.querySelector('i');
    if (icon) icon.className = 'fas fa-spinner fa-pulse';

    var success = 0;
    var failed = 0;
    var errors = [];
    var adminEmails = {};
    usersCache.forEach(function(u) { if (u.role === 'admin') adminEmails[(u.email || '').toLowerCase()] = true; });

    var BATCH_SIZE = 400;
    for (var i = 0; i < toImport.length; i += BATCH_SIZE) {
        var chunk = toImport.slice(i, i + BATCH_SIZE);
        chunk = chunk.filter(function(r) {
            if (adminEmails[r.email]) { console.warn('Skip admin:', r.email); return false; }
            return true;
        });
        if (chunk.length === 0) continue;
        var batch = db.batch();
        chunk.forEach(function(r) {
            var ref = db.collection('allowed_users').doc(r.email);
            var data = { name: r.name, role: 'user', addedBy: currentUser.email };
            if (!r.isUpdate) {
                data.addedAt = firebase.firestore.FieldValue.serverTimestamp();
                data.importedFromExcel = true;
            } else {
                data.updatedAt = firebase.firestore.FieldValue.serverTimestamp();
            }
            if (r.expDate) {
                try { data.expiresAt = firebase.firestore.Timestamp.fromDate(r.expDate); }
                catch(e) { data.expiresAt = null; }
            } else {
                data.expiresAt = null;
            }
            batch.set(ref, data, { merge: true });
        });
        try {
            await batch.commit();
            success += chunk.length;
        } catch(err) {
            failed += chunk.length;
            errors.push(err.message);
            console.error('Batch error:', err);
        }
        var countEl = btn.querySelector('#importCount');
        if (countEl) countEl.textContent = success + '/' + toImport.length;
    }

    btn.disabled = originalDisabled;
    btn.innerHTML = originalHTML;
    importRows = [];
    try { localStorage.removeItem('admin_users_cache'); } catch(e) {}
    toImport.forEach(function(r) { try { localStorage.removeItem('user_cache_' + r.email); } catch(e) {} });

    var msg = '✅ Import hoàn tất!\n\n✓ Thành công: ' + success + '\n' + (failed ? '✗ Thất bại: ' + failed + '\n' : '') + (errors.length ? '\nLỗi:\n' + errors.slice(0, 3).join('\n') : '');
    alert(msg);
    $('importModal').classList.remove('show');
    loadUsers(true);
}

/* ============ DATE HELPERS ============ */
function getExpiryDate(expiresAt) {
    if (!expiresAt) return null;
    try {
        var ea = expiresAt;
        if (typeof ea.toDate === 'function') return ea.toDate();
        if (ea.seconds) return new Date(ea.seconds * 1000);
        return new Date(ea);
    } catch(e) { return null; }
}
function getExpiryTimestamp(expiresAt) {
    var d = getExpiryDate(expiresAt);
    return d ? d.getTime() : null;
}
function formatDate(d) {
    if (!d) return '';
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

/* ============ INIT AUTH UI (chạy sau initApp) ============ */
function initAuthUI() {
    if ($('headerLoginBtn')) $('headerLoginBtn').addEventListener('click', showLoginModal);
    if ($('loginClose')) $('loginClose').addEventListener('click', hideLoginModal);
    if ($('loginModal')) {
        $('loginModal').addEventListener('click', function(e) {
            if (e.target === this) hideLoginModal();
        });
    }
    if ($('googleLoginBtn')) {
        $('googleLoginBtn').addEventListener('click', async function() {
            var provider = new firebase.auth.GoogleAuthProvider();
            provider.setCustomParameters({ prompt: 'select_account' });
            try {
                await auth.signInWithPopup(provider);
                hideLoginModal();
            } catch(e) {
                if (e.code === 'auth/popup-blocked') {
                    await auth.signInWithRedirect(provider);
                } else if (e.code !== 'auth/popup-closed-by-user') {
                    showLoginError('Lỗi: ' + e.message);
                }
            }
        });
    }
    if ($('logoutBtn')) {
        $('logoutBtn').addEventListener('click', function() {
            if (confirm('Đăng xuất?')) {
                try { if (currentUser && currentUser.email) localStorage.removeItem('user_cache_' + currentUser.email); } catch(e) {}
                auth.signOut();
            }
        });
    }
    if ($('userAvatar')) {
        $('userAvatar').addEventListener('click', function(e) {
            e.stopPropagation();
            $('userDropdown').classList.toggle('show');
        });
    }
    document.addEventListener('click', function(e) {
        var dd = $('userDropdown');
        if (dd && !dd.contains(e.target) && $('userAvatar') && !$('userAvatar').contains(e.target)) {
            dd.classList.remove('show');
        }
    });

    if ($('changeNameBtn')) {
        $('changeNameBtn').addEventListener('click', function() {
            $('userDropdown').classList.remove('show');
            if (!currentUser) return;
            editingEmail = currentUser.email;
            $('changeNameCurrent').textContent = currentUser.name;
            $('changeNameEmail').textContent = currentUser.email;
            $('changeNameInput').value = currentUser.name;
            $('changeNameModal').classList.add('show');
            setTimeout(function() { $('changeNameInput').focus(); }, 100);
        });
    }
    if ($('changeNameClose')) $('changeNameClose').addEventListener('click', function() { $('changeNameModal').classList.remove('show'); });
    if ($('changeNameCancel')) $('changeNameCancel').addEventListener('click', function() { $('changeNameModal').classList.remove('show'); });
    if ($('changeNameModal')) {
        $('changeNameModal').addEventListener('click', function(e) {
            if (e.target === this) $('changeNameModal').classList.remove('show');
        });
    }
    if ($('changeNameConfirm')) $('changeNameConfirm').addEventListener('click', doChangeName);

    if ($('editExpiryClose')) $('editExpiryClose').addEventListener('click', function() { $('editExpiryModal').classList.remove('show'); });
    if ($('editExpiryCancel')) $('editExpiryCancel').addEventListener('click', function() { $('editExpiryModal').classList.remove('show'); });
    if ($('editExpiryModal')) {
        $('editExpiryModal').addEventListener('click', function(e) {
            if (e.target === this) $('editExpiryModal').classList.remove('show');
        });
    }
    if ($('editExpiryConfirm')) $('editExpiryConfirm').addEventListener('click', doUpdateExpiry);

    initAdminPanel();
}

setTimeout(function() {
    if (!appInitialized) {
        console.warn('Auth timeout, entering demo mode');
        enterDemoMode();
    }
}, 5000);
"""
    

def build_accounts_js():
    """JS: Firebase auth, login, admin panel, user management, import/export, trial, renewal."""
    return r"""
/* ============ AUTH ============ */
var currentUser = null;
var isDemo = true;
var auth, db;
var usersCache = [];
var lastLoginMap = {};
var importRows = [];
var editingEmail = null;
var editingExpiryEmail = null;
var appInitialized = false;

function isSuperAdmin() {
    if (!currentUser || currentUser.role !== 'admin') return false;
    var email = (currentUser.email || '').toLowerCase().trim();
    return email === SUPER_ADMIN.toLowerCase().trim();
}
function isHiddenAdmin() {
    if (!currentUser || currentUser.role !== 'admin') return false;
    return !isSuperAdmin();
}

try {
    firebase.initializeApp(FIREBASE_CONFIG);
    auth = firebase.auth();
    db = firebase.firestore();
    auth.onAuthStateChanged(handleAuthChange);
} catch(e) {
    console.error('Firebase init error:', e);
    enterDemoMode();
}

async function handleAuthChange(user) {
    if (!user) {
        currentUser = null; isDemo = true;
        applyUserUI(); enterDemoMode();
        return;
    }

    var email = (user.email || '').toLowerCase();
    var cacheKey = 'user_cache_' + email;
    var cached = null;
    try { cached = JSON.parse(localStorage.getItem(cacheKey) || 'null'); } catch(e) {}

    if (cached && cached.expires > Date.now() && cached.data) {
        if (!checkUserExpiration(cached.data)) {
            await auth.signOut();
            try { localStorage.removeItem(cacheKey); } catch(e) {}
            enterDemoMode();
            return;
        }
        currentUser = cached.data;
        isDemo = false;
        applyUserUI();
        logLogin(currentUser);
        if (!appInitialized) { initApp(); appInitialized = true; }
        else { if (typeof refreshApp === 'function') refreshApp(); }
        return;
    }

    try {
        var doc = await db.collection('allowed_users').doc(email).get();

        // ✅ TỰ ĐỘNG ĐĂNG KÝ: user mới → tặng 7 ngày
        if (!doc.exists) {
            if (typeof grantTrialIfNew !== 'function') {
                await auth.signOut();
                showLoginError('Lỗi: Không tải được module trial. Vui lòng tải lại trang.');
                enterDemoMode();
                return;
            }

            var registered = await grantTrialIfNew(user, null);
            if (registered) {
                doc = await db.collection('allowed_users').doc(email).get();
                setTimeout(function() {
                    var trialDate = new Date(Date.now() + (typeof TRIAL_DAYS !== 'undefined' ? TRIAL_DAYS : 7) * 86400000);
                    alert('🎉 Chào mừng bạn đến với Học tiếng Trung!\n\n' +
                          '✅ Bạn được tặng MIỄN PHÍ ' +
                          (typeof TRIAL_DAYS !== 'undefined' ? TRIAL_DAYS : 7) +
                          ' ngày sử dụng.\n\n' +
                          '📅 Hạn dùng: ' + trialDate.toLocaleDateString('vi-VN') + '\n\n' +
                          'Chúc bạn học tốt! 🎓');
                }, 600);
            } else {
                await auth.signOut();
                showLoginError('Tài khoản <b>' + email + '</b> chưa được cấp quyền.');
                enterDemoMode();
                return;
            }
        }

        var data = doc.data() || {};
        var userData = {
            email: email,
            name: data.name || user.displayName || email.split('@')[0],
            role: data.role || 'user',
            photo: user.photoURL || '',
            expiresAt: data.expiresAt || null,
            isTrial: data.isTrial || false
        };

        if (!checkUserExpiration(userData)) {
            await auth.signOut();
            try { localStorage.removeItem(cacheKey); } catch(e) {}
            enterDemoMode();
            return;
        }

        currentUser = userData;

        try {
            localStorage.setItem(cacheKey, JSON.stringify({
                data: currentUser,
                expires: Date.now() + 12 * 60 * 60 * 1000
            }));
        } catch(e) {}

        isDemo = false;
        applyUserUI();
        logLogin(currentUser);
        if (!appInitialized) { initApp(); appInitialized = true; }
        else { if (typeof refreshApp === 'function') refreshApp(); }
    } catch(e) {
        console.error('Auth check error:', e);
        isDemo = true; enterDemoMode();
    }
}

/* ✅ KHÔNG chặn login khi hết hạn */
function checkUserExpiration(userData) {
    return true;
}

function getDaysRemaining(userData) {
    if (!userData || !userData.expiresAt) return null;
    if (userData.role === 'admin') return null;
    var expDate;
    try {
        var ea = userData.expiresAt;
        if (typeof ea.toDate === 'function') expDate = ea.toDate();
        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
        else expDate = new Date(ea);
    } catch(e) { return null; }
    return Math.ceil((expDate.getTime() - Date.now()) / (24 * 60 * 60 * 1000));
}

function enterDemoMode() {
    currentUser = null; isDemo = true;
    applyUserUI();
    if (!appInitialized) { initApp(); appInitialized = true; }
    else { if (typeof refreshApp === 'function') refreshApp(); }
    $('loadingScreen').classList.add('hidden');
    $('stickyTop').style.display = 'block';
    $('fabGroup').style.display = 'flex';
    $('mainContent').style.display = 'block';
}

function applyUserUI() {
    var demoBadge = $('demoBadge');
    var headerLoginBtn = $('headerLoginBtn');
    var userMenu = $('userMenu');

    /* ✅ BANNER HẾT HẠN + NÚT GIA HẠN */
    var expiryBanner = $('expiryBanner');
    if (expiryBanner) {
        if (!isDemo && currentUser && currentUser.role !== 'admin') {
            var daysLeft = getDaysRemaining(currentUser);
            if (daysLeft !== null && daysLeft <= 7) {
                expiryBanner.style.display = 'flex';
                var isExpired = daysLeft <= 0;
                expiryBanner.classList.toggle('urgent', isExpired || daysLeft <= 3);

                var iconWrap = expiryBanner.querySelector('.expiry-banner-icon');
                if (iconWrap) {
                    iconWrap.innerHTML = isExpired
                        ? '<i class="fas fa-exclamation-triangle"></i>'
                        : '<i class="fas fa-hourglass-half"></i>';
                }

                var titleEl = expiryBanner.querySelector('.expiry-banner-text .title');
                if (titleEl) {
                    if (isExpired) titleEl.innerHTML = '❌ Tài khoản đã hết hạn!';
                    else if (daysLeft <= 3) titleEl.innerHTML = '⏰ Sắp hết hạn — còn ' + daysLeft + ' ngày';
                    else titleEl.innerHTML = '⏳ Sắp hết hạn — còn ' + daysLeft + ' ngày';
                }

                var descEl = expiryBanner.querySelector('.expiry-banner-text .desc');
                if (descEl) {
                    var expDateStr = '';
                    try {
                        var ea = currentUser.expiresAt;
                        var expDate;
                        if (typeof ea.toDate === 'function') expDate = ea.toDate();
                        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
                        else expDate = new Date(ea);
                        if (expDate) expDateStr = expDate.toLocaleDateString('vi-VN');
                    } catch(e) {}
                    if (isExpired) {
                        descEl.innerHTML = 'Đã hết hạn vào <b>' + expDateStr + '</b>. Gia hạn ngay để tiếp tục học!';
                    } else {
                        descEl.innerHTML = 'Còn <b>' + daysLeft + ' ngày</b> (đến <b>' + expDateStr + '</b>). Gia hạn để không bị gián đoạn!';
                    }
                }

                var contactBtn = $('expiryContactBtn');
                if (contactBtn) {
                    contactBtn.innerHTML = '<i class="fas fa-crown"></i> Gia hạn ngay';
                    contactBtn.href = '#';
                    contactBtn.onclick = function(e) {
                        e.preventDefault();
                        if (typeof openRenewalModal === 'function') openRenewalModal();
                        else alert('Vui lòng tải lại trang để dùng tính năng gia hạn.');
                    };
                }
            } else {
                expiryBanner.style.display = 'none';
            }
        } else {
            expiryBanner.style.display = 'none';
        }
    }

    var renewBtn = $('dropdownRenewBtn');
    if (renewBtn) {
        renewBtn.style.display = (!isDemo && currentUser && currentUser.role !== 'admin') ? 'flex' : 'none';
    }

    if (isDemo) {
        demoBadge.style.display = 'flex';
        headerLoginBtn.style.display = 'flex';
        userMenu.style.display = 'none';
        $('demoBanner').style.display = 'flex';
        var ud0 = $('userDetails');
        if (ud0) ud0.style.display = 'none';
    } else {
        demoBadge.style.display = 'none';
        headerLoginBtn.style.display = 'none';
        userMenu.style.display = 'block';
        $('demoBanner').style.display = 'none';

        $('userName').textContent = currentUser.name;
        $('userEmail').textContent = currentUser.email;
        $('userRole').textContent = currentUser.role;
        $('userRole').className = 'role' + (currentUser.role === 'admin' ? ' admin' : '');
        $('openAdminBtn').style.display = currentUser.role === 'admin' ? 'flex' : 'none';

        var avatar = $('userAvatar');
        if (currentUser.photo) avatar.src = currentUser.photo;
        else {
            avatar.src = 'data:image/svg+xml;utf8,' + encodeURIComponent(
                '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">' +
                '<rect fill="#2563eb" width="100" height="100"/>' +
                '<text x="50" y="65" font-size="45" fill="#fff" text-anchor="middle" font-family="sans-serif" font-weight="bold">' +
                currentUser.name.charAt(0).toUpperCase() + '</text></svg>'
            );
        }

        updateUserDetails();
    }

    var hskChip = $('hskChip');
    var subjectChip = $('subjectChip');
    if (isDemo) {
        hskChip.classList.add('demo-limited');
        subjectChip.classList.add('demo-limited');
    } else {
        hskChip.classList.remove('demo-limited');
        subjectChip.classList.remove('demo-limited');
    }

    var zaloBtn = $('zaloBtn');
    if (zaloBtn) {
        if (isDemo) zaloBtn.classList.remove('compact');
        else zaloBtn.classList.add('compact');
    }

    $('demoLimitText').textContent = DEMO_LIMIT;
    $('demoDailyText').textContent = DEMO_DAILY_LIMIT;
    var hskMaxEl1 = $('demoHskMaxText');
    var hskMaxEl2 = $('demoHskMaxText2');
    if (hskMaxEl1) hskMaxEl1.textContent = DEMO_HSK_MAX;
    if (hskMaxEl2) hskMaxEl2.textContent = DEMO_HSK_MAX;
    if (typeof updateDemoRemaining === 'function') updateDemoRemaining();

    var toggleFocusBtn = $('toggleFocusBtn');
    if (toggleFocusBtn) {
        toggleFocusBtn.style.display = isDemo ? 'none' : 'flex';
    }

    if (isDemo) {
        document.body.classList.remove('hide-floating');
    }
}

function updateUserDetails() {
    var detailsEl = $('userDetails');
    if (!detailsEl) return;
    if (isDemo || !currentUser) { detailsEl.style.display = 'none'; return; }
    detailsEl.style.display = 'flex';

    var expiryValue = $('expiryValue');
    var expirySub = $('expirySub');
    var expiryIcon = $('expiryIcon');
    var expiryIconWrap = $('expiryIconWrap');
    var progressWrap = $('expiryProgressWrap');
    var progressBar = $('expiryProgressBar');
    if (!expiryValue || !expirySub) return;

    if (currentUser.role === 'admin') {
        expiryValue.textContent = 'Vĩnh viễn';
        expiryValue.className = 'detail-value permanent';
        expirySub.textContent = 'Tài khoản quản trị viên';
        expiryIcon.className = 'fas fa-infinity';
        expiryIconWrap.className = 'detail-icon permanent';
        if (progressWrap) progressWrap.style.display = 'none';
        return;
    }

    if (!currentUser.expiresAt) {
        expiryValue.textContent = 'Vĩnh viễn';
        expiryValue.className = 'detail-value permanent';
        expirySub.textContent = 'Không giới hạn thời gian';
        expiryIcon.className = 'fas fa-infinity';
        expiryIconWrap.className = 'detail-icon permanent';
        if (progressWrap) progressWrap.style.display = 'none';
        return;
    }

    var expDate;
    try {
        var ea = currentUser.expiresAt;
        if (typeof ea.toDate === 'function') expDate = ea.toDate();
        else if (ea.seconds) expDate = new Date(ea.seconds * 1000);
        else expDate = new Date(ea);
    } catch(e) {
        expiryValue.textContent = '-'; expirySub.textContent = ''; return;
    }
    if (!expDate || isNaN(expDate.getTime())) {
        expiryValue.textContent = '-'; expirySub.textContent = ''; return;
    }

    var now = Date.now();
    var expTime = expDate.getTime();
    var daysLeft = Math.ceil((expTime - now) / (24 * 60 * 60 * 1000));
    var dateStr = expDate.toLocaleDateString('vi-VN');

    expiryValue.className = 'detail-value';
    expiryIconWrap.className = 'detail-icon';

    if (daysLeft < 0) {
        expiryValue.textContent = 'Đã hết hạn';
        expiryValue.classList.add('expired');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Đã hết hạn ' + Math.abs(daysLeft) + ' ngày trước';
        expiryIcon.className = 'fas fa-calendar-times';
        expiryIconWrap.classList.add('expired');
        if (progressWrap) progressWrap.style.display = 'none';
    } else if (daysLeft === 0) {
        expiryValue.textContent = 'Hết hạn hôm nay';
        expiryValue.classList.add('urgent');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Vui lòng liên hệ Admin để gia hạn!';
        expiryIcon.className = 'fas fa-exclamation-circle';
        expiryIconWrap.classList.add('urgent');
        if (progressWrap) progressWrap.style.display = 'none';
    } else if (daysLeft <= 3) {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('urgent');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b><br>Sắp hết hạn, vui lòng gia hạn!';
        expiryIcon.className = 'fas fa-exclamation-circle';
        expiryIconWrap.classList.add('urgent');
        var total3 = 30 * 24 * 60 * 60 * 1000;
        var pct3 = Math.max(0, Math.min(100, ((total3 - (expTime - now)) / total3) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar urgent';
            progressBar.style.width = pct3 + '%';
        }
    } else if (daysLeft <= 7) {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('warn');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b>';
        expiryIcon.className = 'fas fa-clock';
        expiryIconWrap.classList.add('warn');
        var total7 = 30 * 24 * 60 * 60 * 1000;
        var pct7 = Math.max(0, Math.min(100, ((total7 - (expTime - now)) / total7) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar warn';
            progressBar.style.width = pct7 + '%';
        }
    } else {
        expiryValue.textContent = 'Còn ' + daysLeft + ' ngày';
        expiryValue.classList.add('ok');
        expirySub.innerHTML = 'Ngày hết hạn: <b>' + dateStr + '</b>';
        expiryIcon.className = 'fas fa-calendar-check';
        expiryIconWrap.classList.add('ok');
        var totalOk = 30 * 24 * 60 * 60 * 1000;
        var pctOk = Math.max(0, Math.min(100, ((totalOk - (expTime - now)) / totalOk) * 100));
        if (progressWrap) {
            progressWrap.style.display = 'block';
            progressBar.className = 'progress-bar ok';
            progressBar.style.width = pctOk + '%';
        }
    }
}

/* ============ LOGIN UI ============ */
window.showLoginModal = function() {
    $('loginModal').classList.add('show');
    $('loginError').classList.remove('show');
};
function hideLoginModal() { $('loginModal').classList.remove('show'); }

function showLoginError(msg) {
    var el = $('loginError');
    el.innerHTML = '<i class="fas fa-exclamation-triangle"></i> ' + msg;
    el.classList.add('show');
}

function logLogin(u) {
    try {
        var today = new Date().toDateString();
        var logKey = 'login_log_' + u.email;
        if (localStorage.getItem(logKey) === today) return;
        db.collection('login_logs').add({
            email: u.email, name: u.name, role: u.role,
            time: firebase.firestore.FieldValue.serverTimestamp(),
            userAgent: navigator.userAgent.substring(0, 100)
        }).then(function() {
            try { localStorage.setItem(logKey, today); } catch(e) {}
        }).catch(function() {});
    } catch(e) {}
}
