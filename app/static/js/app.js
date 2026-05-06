$(document).ready(function () {

    // ── Config ────────────────────────────────────────────────
    const DEFAULT_VISIBLE_COLUMNS = [
        "S.No", "Emp Code", "Emp Name", "Status", "Skills", "Experience",
        "Fresher/Lateral", "Offshore/Onsite", "Designation", "Grade",
        "Gender", "Sub Practice", "Billable/Non Billable", "Projects",
        "Customer Name", "Customer interview happened(Yes/No)",
        "Customer Selected(Yes/No)", "Comments"
    ];

    const STORAGE_KEY = "hcr_visible_columns";

    function loadVisibleColumns() {
        try {
            const saved = localStorage.getItem(STORAGE_KEY);
            if (saved) {
                const parsed = JSON.parse(saved);
                if (Array.isArray(parsed) && parsed.length > 0) return parsed;
            }
        } catch (e) { /* ignore */ }
        return DEFAULT_VISIBLE_COLUMNS.slice();
    }

    let visibleColumns = loadVisibleColumns();

    const DEMAND_VISIBLE_COLUMNS = [
        "Requisition ID", "Yrs of Exp", "Skillset", "Customer Name",
        "Demand Status", "Fulfillment Type", "Mapped Emp Code",
        "Mapped Emp Name", "Mapping Date", "Notes"
    ];

    const EDITABLE_COLUMNS = [
        "Emp Name", "Status", "Skills", "Fresher/Lateral", "Offshore/Onsite",
        "Designation", "Grade", "First Line Manager", "Skip Level Manager",
        "Sub Practice", "Remarks", "Billable/Non Billable", "Projects",
        "Remarks2", "Customer Name", "Customer interview happened(Yes/No)",
        "Customer Selected(Yes/No)", "Comments"
    ];

    const READONLY_COLUMNS = ["S.No", "Emp Code", "row_index"];

    const ALL_COLUMNS = [
        "S.No", "Division", "Emp Code", "Emp Name", "Status", "LWD", "Skills",
        "Fresher/Lateral", "Offshore/Onsite", "Experience", "Designation",
        "Grade", "DOJ", "Gender", "First Line Manager", "Skip Level Manager",
        "Company Email", "Sub Practice", "Remarks", "Empower SL",
        "Correction to LM", "Billable/Non Billable", "Billable Till Date",
        "Projects", "Remarks2", "Customer Name",
        "Customer interview happened(Yes/No)", "Customer Selected(Yes/No)",
        "Comments"
    ];

    // ── State ─────────────────────────────────────────────────
    let currentPage = 1;
    let perPage = 25;
    let currentData = [];
    let sortColumn = null;
    let sortAsc = true;
    let activeKpi = "all";

    let demandPage = 1;
    let demandPerPage = 25;
    let demandData = [];
    let activeSuggestDemandRow = null;
    let activeDfKpi = "all";

    // ── Toast Config ──────────────────────────────────────────
    toastr.options = {
        closeButton: true,
        progressBar: true,
        positionClass: "toast-top-right",
        timeOut: 3000,
    };

    // ── Init ──────────────────────────────────────────────────
    loadFilters();
    loadData();
    loadSummary();
    loadDemandSummary();
    renderColumnChooser();
    refreshNotifAwaitingBadge();

    // ── Column Chooser ─────────────────────────────────────────
    function renderColumnChooser() {
        const $list = $("#columnChooserList");
        $list.empty();
        ALL_COLUMNS.forEach(function (col) {
            const checked = visibleColumns.includes(col) ? "checked" : "";
            $list.append(
                `<label class="col-chooser-item d-flex align-items-center gap-2 px-2 py-1 rounded">
                    <input type="checkbox" class="form-check-input col-toggle m-0" value="${col}" ${checked}>
                    <span class="col-chooser-label">${col}</span>
                </label>`
            );
        });
    }

    $(document).on("change", ".col-toggle", function () {
        const col = $(this).val();
        if ($(this).is(":checked")) {
            if (!visibleColumns.includes(col)) {
                const refIndex = ALL_COLUMNS.indexOf(col);
                let inserted = false;
                for (let i = 0; i < visibleColumns.length; i++) {
                    if (ALL_COLUMNS.indexOf(visibleColumns[i]) > refIndex) {
                        visibleColumns.splice(i, 0, col);
                        inserted = true;
                        break;
                    }
                }
                if (!inserted) visibleColumns.push(col);
            }
        } else {
            visibleColumns = visibleColumns.filter(function (c) { return c !== col; });
        }
        localStorage.setItem(STORAGE_KEY, JSON.stringify(visibleColumns));
        renderTableHeader();
        renderTableBody(currentData);
    });

    $("#btnResetColumns").click(function () {
        visibleColumns = DEFAULT_VISIBLE_COLUMNS.slice();
        localStorage.removeItem(STORAGE_KEY);
        renderColumnChooser();
        renderTableHeader();
        renderTableBody(currentData);
    });

    // ── Filter Dropdowns ──────────────────────────────────────
    function loadFilters() {
        $.getJSON("/api/filters", function (resp) {
            const $sp = $("#filterSubPractice");
            const $bl = $("#filterBillable");
            const $pj = $("#filterProject");

            resp.sub_practices.forEach(function (v) {
                $sp.append(`<option value="${v}">${v}</option>`);
            });
            resp.billable_options.forEach(function (v) {
                $bl.append(`<option value="${v}">${v}</option>`);
            });
            (resp.projects || []).forEach(function (v) {
                $pj.append(`<option value="${v}">${v}</option>`);
            });

            $sp.select2({ theme: "bootstrap-5", allowClear: true, placeholder: "All Sub Practices" });
            $bl.select2({ theme: "bootstrap-5", allowClear: true, placeholder: "All Status" });
            $pj.select2({ theme: "bootstrap-5", allowClear: true, placeholder: "All Projects" });
        });
    }

    // ── Summary Cards + Charts ────────────────────────────────
    function loadSummary() {
        const params = getFilterParams();
        $.getJSON("/api/summary", params, function (resp) {
            animateCounter("#totalCount", resp.total);
            animateCounter("#dataCount", resp.data_count);
            animateCounter("#aiCount", resp.ai_count);
            animateCounter("#biCount", resp.bi_count);
            animateCounter("#coreCount", resp.core_count);
            animateCounter("#notConfirmedCount", resp.not_confirmed_count);
            animateCounter("#blTotalCount", resp.total);
            animateCounter("#billableCount", resp.billable);
            animateCounter("#nonBillableCount", resp.non_billable);
            animateCounter("#blockedCount", resp.blocked);
            animateCounter("#proposedCount", resp.proposed);
            animateCounter("#internalProjectCount", resp.internal_project_count);
            animateCounter("#solutionOfferingCount", resp.solution_offering_count);
            animateCounter("#otherCount", resp.other);

            $("#summaryTotal").text(resp.total);
            $("#summaryBillable").text(resp.billable);
            $("#summaryNonBillable").text(resp.non_billable);
            $("#summaryBlocked").text(resp.blocked);
        });
    }

    function loadDemandSummary() {
        $.getJSON("/api/demand-fulfillment/summary", function (resp) {
            animateCounter("#dfTotalDemands", resp.total_demands);
            animateCounter("#dfOpenDemands", resp.open_demands);
            animateCounter("#dfInternallyFulfilled", resp.internally_fulfilled);
            animateCounter("#dfExternallyFulfilled", resp.externally_fulfilled);
            animateCounter("#dfExternalRaised", resp.external_raised);

            $("#summaryOpenDemands").text(resp.open_demands);
        });
    }

    function animateCounter(selector, target) {
        const $el = $(selector);
        if (!$el.length) return;
        const start = parseInt($el.text()) || 0;
        if (start === target) { $el.text(target); return; }
        $({ val: start }).animate({ val: target }, {
            duration: 500,
            step: function () { $el.text(Math.floor(this.val)); },
            complete: function () { $el.text(target); }
        });
    }

    // ── Head Count Data Table ─────────────────────────────────
    function loadData() {
        const params = getFilterParams();
        params.page = currentPage;
        params.per_page = perPage;

        $.getJSON("/api/headcount", params, function (resp) {
            currentData = resp.data;
            renderTableHeader();
            renderTableBody(resp.data);
            renderPagination(resp.total, resp.page, resp.total_pages);
            $("#filteredCount").text(resp.total + " records");
        }).fail(function (xhr) {
            const err = xhr.responseJSON ? xhr.responseJSON.error : "Failed to load data";
            toastr.error(err);
            $("#tableBody").html(
                `<tr><td colspan="20" class="text-center py-4 text-danger">${err}</td></tr>`
            );
        });
    }

    function renderTableHeader() {
        let html = '<th class="text-center" style="width: 50px;">#</th>';
        visibleColumns.forEach(function (col) {
            const isActive = sortColumn === col;
            const icon = isActive ? (sortAsc ? "bi-sort-up" : "bi-sort-down") : "bi-chevron-expand";
            const cls = isActive ? "sort-active" : "";
            html += `<th class="${cls}" data-col="${col}">${col} <i class="bi ${icon} sort-icon"></i></th>`;
        });
        html += '<th class="text-center" style="width: 70px;">Action</th>';
        $("#tableHeader").html(html);
    }

    function renderTableBody(data) {
        if (!data.length) {
            $("#tableBody").html(
                '<tr><td colspan="20" class="text-center py-4 text-muted">No records found</td></tr>'
            );
            return;
        }

        let html = "";
        data.forEach(function (row, i) {
            html += `<tr data-row-index="${row.row_index}">`;
            html += `<td class="text-center text-muted small">${(currentPage - 1) * perPage + i + 1}</td>`;
            visibleColumns.forEach(function (col) {
                const val = row[col] != null ? row[col] : "";
                let display = escapeHtml(String(val));

                if (col === "Billable/Non Billable") {
                    const cls = val === "Billable" ? "badge-billable"
                              : val === "Non-Billable" ? "badge-nonbillable"
                              : val === "Blocked" ? "badge-blocked"
                              : val === "Proposed" ? "badge-proposed"
                              : "badge-other";
                    display = `<span class="badge ${cls}">${display}</span>`;
                }
                if (col === "Sub Practice") {
                    display = `<span class="badge bg-primary">${display}</span>`;
                }
                if (col === "Customer interview happened(Yes/No)" || col === "Customer Selected(Yes/No)") {
                    if (val === "Yes") display = `<span class="badge bg-success">Yes</span>`;
                    else if (val === "No") display = `<span class="badge bg-danger">No</span>`;
                }
                html += `<td title="${escapeHtml(String(val))}">${display}</td>`;
            });
            html += `<td class="text-center text-nowrap">
                ${renderNotifyButton(row)}
                <button class="btn btn-outline-primary btn-sm btn-edit-row" data-row='${JSON.stringify(row)}'>
                    <i class="bi bi-pencil"></i>
                </button>
                <button class="btn btn-outline-danger btn-sm btn-delete-row" data-row-index="${row.row_index}" data-emp-name="${escapeHtml(String(row["Emp Name"] || ""))}" data-emp-code="${escapeHtml(String(row["Emp Code"] || ""))}">
                    <i class="bi bi-trash"></i>
                </button>
            </td>`;
            html += "</tr>";
        });
        $("#tableBody").html(html);
    }

    function renderNotifyButton(row) {
        const blStatus = (row["Billable/Non Billable"] || "").trim();
        const isProposed = blStatus === "Proposed";
        const isBlocked = blStatus === "Blocked";
        if (!isProposed && !isBlocked) return "";

        const n = row._notification;
        const empName = escapeHtml(String(row["Emp Name"] || ""));
        const empCode = escapeHtml(String(row["Emp Code"] || ""));
        const empRow = row.row_index;

        if (n && n.status === "awaiting_reply") {
            const days = relativeDays(n.sent_at);
            return `<span class="badge badge-notif-awaiting me-1" title="Sent ${escapeHtml(n.sent_at || '')} • ${n.reminder_count || 0} reminder(s)">
                        <i class="bi bi-hourglass-split"></i> ${days} • ${n.reminder_count || 0}r
                    </span>`;
        }

        const badgeHtml = n ? renderNotifBadge(n.status) + " " : "";

        if (isBlocked) {
            return `${badgeHtml}<button class="btn btn-warning btn-sm btn-check-allocation"
                        data-emp-row="${empRow}"
                        data-emp-name="${empName}"
                        data-emp-code="${empCode}"
                        title="Check Allocation Status">
                    <i class="bi bi-question-circle-fill"></i>
                </button>`;
        }

        return `${badgeHtml}<button class="btn btn-success btn-sm btn-notify-manager"
                    data-emp-row="${empRow}"
                    data-emp-name="${empName}"
                    data-emp-code="${empCode}"
                    title="Notify First Line Manager">
                <i class="bi bi-send-fill"></i>
            </button>`;
    }

    function renderNotifBadge(status) {
        const map = {
            "awaiting_reply": ["badge-notif-awaiting", "Awaiting"],
            "approved": ["badge-notif-approved", "Approved"],
            "rejected": ["badge-notif-rejected", "Rejected"],
            "no_response": ["badge-notif-noresp", "No Response"],
            "cancelled": ["badge-notif-cancel", "Cancelled"],
        };
        const [cls, label] = map[status] || ["badge bg-secondary", status];
        return `<span class="badge ${cls}">${label}</span>`;
    }

    function relativeDays(isoStr) {
        if (!isoStr) return "?";
        const sent = new Date(isoStr);
        if (isNaN(sent)) return isoStr;
        const diffMs = Date.now() - sent.getTime();
        const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
        if (diffDays <= 0) return "today";
        if (diffDays === 1) return "1d ago";
        return `${diffDays}d ago`;
    }

    // ── Pagination (reusable) ─────────────────────────────────
    function buildPaginationHtml(total, page, totalPages, pp) {
        const start = Math.min((page - 1) * pp + 1, total);
        const end = Math.min(page * pp, total);
        const info = `Showing ${start}-${end} of ${total}`;

        let html = "";
        html += `<li class="page-item ${page <= 1 ? 'disabled' : ''}">
            <a class="page-link" href="#" data-page="${page - 1}"><i class="bi bi-chevron-left"></i></a></li>`;

        const maxVisible = 5;
        let startPage = Math.max(1, page - Math.floor(maxVisible / 2));
        let endPage = Math.min(totalPages, startPage + maxVisible - 1);
        if (endPage - startPage < maxVisible - 1) startPage = Math.max(1, endPage - maxVisible + 1);

        if (startPage > 1) {
            html += `<li class="page-item"><a class="page-link" href="#" data-page="1">1</a></li>`;
            if (startPage > 2) html += `<li class="page-item disabled"><span class="page-link">...</span></li>`;
        }
        for (let p = startPage; p <= endPage; p++) {
            html += `<li class="page-item ${p === page ? 'active' : ''}"><a class="page-link" href="#" data-page="${p}">${p}</a></li>`;
        }
        if (endPage < totalPages) {
            if (endPage < totalPages - 1) html += `<li class="page-item disabled"><span class="page-link">...</span></li>`;
            html += `<li class="page-item"><a class="page-link" href="#" data-page="${totalPages}">${totalPages}</a></li>`;
        }

        html += `<li class="page-item ${page >= totalPages ? 'disabled' : ''}">
            <a class="page-link" href="#" data-page="${page + 1}"><i class="bi bi-chevron-right"></i></a></li>`;

        return { info, controls: html };
    }

    function renderPagination(total, page, totalPages) {
        const pg = buildPaginationHtml(total, page, totalPages, perPage);
        $("#paginationInfo").text(pg.info);
        $("#paginationControls").html(pg.controls);
    }

    // ── Demand Fulfillment Table ──────────────────────────────
    function loadDemandData() {
        const params = {
            demand_status: $("#filterDemandStatus").val() || "All",
            search: $("#demandSearchInput").val() || "",
            df_kpi: activeDfKpi,
            page: demandPage,
            per_page: demandPerPage,
        };

        $.getJSON("/api/demand-requisitions", params, function (resp) {
            demandData = resp.data;
            renderDemandTable(resp.data);
            const pg = buildPaginationHtml(resp.total, resp.page, resp.total_pages, demandPerPage);
            $("#demandPaginationInfo").text(pg.info);
            $("#demandPaginationControls").html(pg.controls);
            $("#demandFilteredCount").text(resp.total + " records");
        }).fail(function (xhr) {
            const err = xhr.responseJSON ? xhr.responseJSON.error : "Failed to load demand data";
            toastr.error(err);
            $("#demandTableBody").html(
                `<tr><td colspan="12" class="text-center py-4 text-danger">${err}</td></tr>`
            );
        });
    }

    function renderDemandTable(data) {
        let headerHtml = '<th class="text-center" style="width:40px">#</th>';
        DEMAND_VISIBLE_COLUMNS.forEach(function (col) {
            headerHtml += `<th>${col}</th>`;
        });
        headerHtml += '<th class="text-center" style="min-width:180px">Actions</th>';
        $("#demandTableHeader").html(headerHtml);

        if (!data.length) {
            $("#demandTableBody").html(
                '<tr><td colspan="12" class="text-center py-4 text-muted">No demand requisitions found</td></tr>'
            );
            return;
        }

        let html = "";
        data.forEach(function (row, i) {
            html += `<tr>`;
            html += `<td class="text-center text-muted small">${(demandPage - 1) * demandPerPage + i + 1}</td>`;
            DEMAND_VISIBLE_COLUMNS.forEach(function (col) {
                const val = row[col] != null ? row[col] : "";
                let display = escapeHtml(String(val));

                if (col === "Demand Status") {
                    const statusMap = {
                        "Open": "bg-warning text-dark",
                        "In Progress": "bg-info text-dark",
                        "Fulfilled": "bg-success",
                        "External": "bg-danger",
                        "On Hold": "bg-secondary",
                        "Closed": "bg-dark"
                    };
                    const cls = statusMap[val] || "bg-secondary";
                    display = `<span class="badge ${cls}">${display}</span>`;
                }
                if (col === "Fulfillment Type") {
                    if (val === "Internal") display = `<span class="badge bg-success">Internal</span>`;
                    else if (val === "External") display = `<span class="badge bg-danger">External</span>`;
                }
                html += `<td title="${escapeHtml(String(val))}">${display}</td>`;
            });

            const isMapped = row["Mapped Emp Code"] && row["Mapped Emp Code"] !== "";
            const demandStatus = row["Demand Status"] || "";
            const isOpen = !demandStatus || demandStatus === "Open" || demandStatus === "In Progress";
            const isFulfilled = demandStatus === "Fulfilled";
            const isExternal = demandStatus === "External";

            html += `<td class="text-center text-nowrap">`;
            html += `<div class="d-flex align-items-center justify-content-center gap-1 flex-wrap">`;

            if (!isMapped && isOpen && !isExternal) {
                html += `<button class="btn btn-sm btn-outline-success btn-find-matches" data-row='${JSON.stringify(row)}' title="Find matching bench employees">
                    <i class="bi bi-search"></i>
                </button>`;
            } else if (isMapped && !isFulfilled && !isExternal) {
                html += `<button class="btn btn-sm btn-success fw-semibold btn-confirm-demand" data-row-index="${row.row_index}" title="Confirm fulfillment — employee becomes Billable">
                    <i class="bi bi-check2-circle me-1"></i>Confirm
                </button>`;
            } else if (isMapped && isFulfilled) {
                html += `<span class="badge bg-success" style="font-size:.72rem"><i class="bi bi-check-circle-fill me-1"></i>Fulfilled</span>`;
            } else if (isExternal) {
                html += `<span class="badge bg-danger me-1" style="font-size:.72rem"><i class="bi bi-box-arrow-up-right me-1"></i>External</span>`;
                html += `<button class="btn btn-sm btn-warning fw-semibold btn-add-external" data-row='${JSON.stringify(row)}' title="Add external employee and fulfill demand">
                    <i class="bi bi-person-plus me-1"></i>Add External
                </button>`;
            }

            html += `<button class="btn btn-sm btn-outline-danger btn-delete-demand" data-row-index="${row.row_index}" data-req-id="${escapeHtml(String(row["Requisition ID"] || ""))}" data-is-mapped="${isMapped ? '1' : '0'}" title="Delete demand">
                <i class="bi bi-trash"></i>
            </button>`;
            html += `</div>`;
            html += `</td></tr>`;
        });
        $("#demandTableBody").html(html);
    }

    // ── Suggestions Modal ─────────────────────────────────────
    function openSuggestionsModal(demandRow) {
        activeSuggestDemandRow = demandRow;
        const reqId = demandRow["Requisition ID"] || "N/A";
        const skillset = demandRow["Skillset"] || "N/A";
        const exp = demandRow["Yrs of Exp"] || "Any";
        const customer = demandRow["Customer Name"] || "N/A";

        $("#suggestReqId").text(reqId);
        $("#suggestSkillset").text(skillset);
        $("#suggestExp").text(exp);
        $("#suggestCustomer").text(customer);
        $("#suggestMatchCount").text("Searching...");
        $("#suggestionsBody").html(
            '<tr><td colspan="8" class="text-center py-4"><div class="spinner-border spinner-border-sm text-primary"></div> Finding matches...</td></tr>'
        );

        new bootstrap.Modal(document.getElementById("suggestionsModal")).show();

        $.getJSON(`/api/demand-fulfillment/suggestions/${demandRow.row_index}`, function (resp) {
            $("#suggestMatchCount").text(resp.total_matches + " matches found");
            renderSuggestions(resp.suggestions);
        }).fail(function (xhr) {
            const err = xhr.responseJSON ? xhr.responseJSON.error : "Failed to find matches";
            $("#suggestionsBody").html(
                `<tr><td colspan="8" class="text-center py-4 text-danger">${err}</td></tr>`
            );
            $("#suggestMatchCount").text("Error");
        });
    }

    function renderSuggestions(suggestions) {
        if (!suggestions.length) {
            $("#suggestionsBody").html(
                '<tr><td colspan="8" class="text-center py-4">' +
                    '<p class="text-muted mb-3">No matching bench employees found.</p>' +
                    '<button class="btn btn-danger btn-mark-external" data-row-index="' + activeSuggestDemandRow.row_index + '">' +
                        '<i class="bi bi-box-arrow-up-right me-1"></i>Mark as External' +
                    '</button>' +
                '</td></tr>'
            );
            return;
        }

        let html = "";
        suggestions.forEach(function (emp) {
            const matchBadges = emp.matched_skills.map(function (s) {
                return `<span class="badge bg-success-subtle text-success me-1" style="font-size:.65rem">${escapeHtml(s)}</span>`;
            }).join("");
            const parsedExp = emp.parsed_experience;
            const expDisplay = parsedExp != null ? parsedExp + " yrs" : "N/A";
            const expClass = emp.exp_match ? "text-success" : "text-warning";

            html += `<tr>
                <td class="small">${escapeHtml(String(emp["Emp Code"] || ""))}</td>
                <td class="small fw-semibold">${escapeHtml(String(emp["Emp Name"] || ""))}</td>
                <td class="small">${escapeHtml(String(emp["Skills"] || ""))}</td>
                <td class="small ${expClass}">${expDisplay} ${emp.exp_match ? '<i class="bi bi-check-circle-fill"></i>' : ''}</td>
                <td class="small"><span class="badge bg-primary">${escapeHtml(String(emp["Sub Practice"] || ""))}</span></td>
                <td class="small"><span class="badge badge-other">${escapeHtml(String(emp["Billable/Non Billable"] || ""))}</span></td>
                <td>${matchBadges || '<span class="text-muted small">None</span>'}</td>
                <td class="text-center">
                    <button class="btn btn-sm btn-success btn-map-employee"
                            data-emp-row="${emp.row_index}"
                            data-emp-name="${escapeHtml(String(emp["Emp Name"] || ""))}"
                            data-emp-code="${escapeHtml(String(emp["Emp Code"] || ""))}">
                        <i class="bi bi-person-check me-1"></i>Map
                    </button>
                </td>
            </tr>`;
        });
        $("#suggestionsBody").html(html);
    }

    function mapEmployee(empRowIndex, empName, empCode) {
        if (!activeSuggestDemandRow) return;

        if (!confirm(`Map ${empName} (${empCode}) to ${activeSuggestDemandRow["Requisition ID"]}?\n\nThis will:\n- Set employee status to "Proposed"\n- Mark demand as "Internal" fulfillment`)) {
            return;
        }

        showLoading();
        $.ajax({
            url: "/api/demand-fulfillment/map",
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify({
                demand_row_index: activeSuggestDemandRow.row_index,
                emp_row_index: empRowIndex,
            }),
            success: function (resp) {
                hideLoading();
                bootstrap.Modal.getInstance(document.getElementById("suggestionsModal")).hide();
                toastr.success(resp.message || "Employee mapped successfully!");
                loadDemandData();
                loadDemandSummary();
                loadData();
                loadSummary();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Mapping failed";
                toastr.error(err);
            }
        });
    }

    function markAsExternal(demandRowIndex) {
        if (!confirm("Mark this demand as External requisition?\n\nThis means no internal bench employee will be mapped.")) return;

        showLoading();
        $.ajax({
            url: `/api/demand-requisition/${demandRowIndex}`,
            method: "PUT",
            contentType: "application/json",
            data: JSON.stringify({ "Fulfillment Type": "External", "Demand Status": "External" }),
            success: function () {
                hideLoading();
                toastr.success("Marked as External requisition");
                var modal = bootstrap.Modal.getInstance(document.getElementById("suggestionsModal"));
                if (modal) modal.hide();
                loadDemandData();
                loadDemandSummary();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Update failed";
                toastr.error(err);
            }
        });
    }

    // ── Edit Modal ────────────────────────────────────────────
    function openEditModal(rowData) {
        let html = '<div class="row g-3">';
        ALL_COLUMNS.forEach(function (col) {
            const val = rowData[col] != null ? rowData[col] : "";
            const isReadonly = READONLY_COLUMNS.includes(col);
            const isEditable = EDITABLE_COLUMNS.includes(col);
            const colSize = col === "Remarks" || col === "Remarks2" ? "12" : "6";

            html += `<div class="col-md-${colSize}">`;
            html += `<label class="form-label">${col}</label>`;

            if (isReadonly) {
                html += `<input type="text" class="form-control" value="${escapeHtml(String(val))}" readonly>`;
            } else if (col === "Sub Practice") {
                html += `<select class="form-select edit-field" data-col="${col}">`;
                ["AI", "BI", "DE", "Core", "Not Confirmed"].forEach(function (opt) {
                    const sel = opt === val ? "selected" : "";
                    html += `<option value="${opt}" ${sel}>${opt}</option>`;
                });
                html += `</select>`;
            } else if (col === "Billable/Non Billable") {
                html += `<select class="form-select edit-field" data-col="${col}">`;
                ["Billable", "Non-Billable", "Internal project", "Proposed", "Shadow- Client Project",
                 "Solution Offerings", "Core", "Resigned", "Serving Notice Period",
                 "Partially Billable", "Blocked", "Not Confirmed", "Internal POC",
                 "Pending Allocation", "Not a Data AI/ML"].forEach(function (opt) {
                    const sel = opt === val ? "selected" : "";
                    html += `<option value="${opt}" ${sel}>${opt}</option>`;
                });
                html += `</select>`;
            } else if (col === "Customer interview happened(Yes/No)" || col === "Customer Selected(Yes/No)") {
                html += `<select class="form-select edit-field" data-col="${col}">`;
                ["", "Yes", "No"].forEach(function (opt) {
                    const sel = opt === val ? "selected" : "";
                    const label = opt === "" ? "-- Select --" : opt;
                    html += `<option value="${opt}" ${sel}>${label}</option>`;
                });
                html += `</select>`;
            } else if (col === "Gender") {
                html += `<select class="form-select edit-field" data-col="${col}">`;
                ["Male", "Female", "Other"].forEach(function (opt) {
                    const sel = opt === val ? "selected" : "";
                    html += `<option value="${opt}" ${sel}>${opt}</option>`;
                });
                html += `</select>`;
            } else if (col === "Status") {
                html += `<select class="form-select edit-field" data-col="${col}">`;
                ["Confirmed", "Resigned", "Not Confirmed", "Absconded"].forEach(function (opt) {
                    const sel = opt === val ? "selected" : "";
                    html += `<option value="${opt}" ${sel}>${opt}</option>`;
                });
                html += `</select>`;
            } else if (isEditable) {
                html += `<input type="text" class="form-control edit-field" data-col="${col}" value="${escapeHtml(String(val))}">`;
            } else {
                html += `<input type="text" class="form-control" value="${escapeHtml(String(val))}" readonly>`;
            }
            html += `</div>`;
        });
        html += "</div>";

        $("#editModalBody").html(html);
        $("#editModal").data("rowIndex", rowData.row_index);
        $("#editModal").data("originalData", rowData);
        new bootstrap.Modal(document.getElementById("editModal")).show();
    }

    function saveEdit() {
        const rowIndex = $("#editModal").data("rowIndex");
        const original = $("#editModal").data("originalData");
        const updates = {};

        $(".edit-field").each(function () {
            const col = $(this).data("col");
            const newVal = $(this).val();
            const oldVal = original[col] != null ? String(original[col]) : "";
            if (newVal !== oldVal) {
                updates[col] = newVal;
            }
        });

        if (Object.keys(updates).length === 0) {
            toastr.info("No changes detected");
            return;
        }

        showLoading();
        $.ajax({
            url: `/api/headcount/${rowIndex}`,
            method: "PUT",
            contentType: "application/json",
            data: JSON.stringify(updates),
            success: function (resp) {
                hideLoading();
                bootstrap.Modal.getInstance(document.getElementById("editModal")).hide();
                toastr.success(resp.message || "Saved to SharePoint!");
                loadData();
                loadSummary();
                loadDemandSummary();
                if ($("#pane-demand").hasClass("show")) loadDemandData();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Save failed";
                toastr.error(err);
            }
        });
    }

    // ── Export CSV ─────────────────────────────────────────────
    function exportCSV() {
        const params = getFilterParams();
        params.page = 1;
        params.per_page = 99999;

        toastr.info("Preparing export...");
        $.getJSON("/api/headcount", params, function (resp) {
            if (!resp.data.length) { toastr.warning("No data to export"); return; }

            const cols = ALL_COLUMNS;
            let csv = cols.map(c => `"${c}"`).join(",") + "\n";
            resp.data.forEach(function (row) {
                csv += cols.map(function (c) {
                    const v = row[c] != null ? String(row[c]).replace(/"/g, '""') : "";
                    return `"${v}"`;
                }).join(",") + "\n";
            });

            const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = `headcount_report_${new Date().toISOString().slice(0, 10)}.csv`;
            link.click();
            toastr.success("Export complete!");
        });
    }

    // ── Event Handlers ────────────────────────────────────────

    // KPI card click
    $(document).on("click", ".stat-card-clickable", function () {
        const kpi = $(this).data("kpi");
        $(".stat-card-clickable").removeClass("active");
        $(this).addClass("active");

        if (kpi === activeKpi) {
            activeKpi = "all";
            $(".stat-card-clickable").removeClass("active");
            $('[data-kpi="all"]').addClass("active");
        } else {
            activeKpi = kpi;
        }
        currentPage = 1;
        loadData();
        loadSummary();
    });

    // Demand Fulfillment KPI card click
    $(document).on("click", ".df-card-clickable", function () {
        const kpi = $(this).data("df-kpi");
        $(".df-card-clickable").removeClass("active");
        $(this).addClass("active");

        if (kpi === activeDfKpi) {
            activeDfKpi = "all";
            $(".df-card-clickable").removeClass("active");
            $('[data-df-kpi="all"]').addClass("active");
        } else {
            activeDfKpi = kpi;
        }
        demandPage = 1;
        loadDemandData();

        // Auto-switch to the Demand Fulfillment tab if not already active
        const dfTab = document.querySelector('#tab-demand');
        if (dfTab && !dfTab.classList.contains('active')) {
            new bootstrap.Tab(dfTab).show();
        }
    });

    $("#btnApply").click(function () { currentPage = 1; loadData(); loadSummary(); });

    $("#btnReset").click(function () {
        $("#filterSubPractice").val("All").trigger("change");
        $("#filterBillable").val("All").trigger("change");
        $("#filterProject").val("All").trigger("change");
        $("#searchInput").val("");
        activeKpi = "all";
        $(".stat-card-clickable").removeClass("active");
        $('[data-kpi="all"]').addClass("active");
        currentPage = 1;
        loadData();
        loadSummary();
    });

    $("#searchInput").on("keydown", function (e) {
        if (e.key === "Enter") { currentPage = 1; loadData(); }
    });

    $("#perPageSelect").change(function () {
        perPage = parseInt($(this).val());
        currentPage = 1;
        loadData();
    });

    $(document).on("click", "#paginationControls .page-link", function (e) {
        e.preventDefault();
        const p = parseInt($(this).data("page"));
        if (!isNaN(p) && p >= 1) { currentPage = p; loadData(); }
    });

    $(document).on("click", ".btn-edit-row", function () {
        openEditModal($(this).data("row"));
    });

    $(document).on("click", ".btn-delete-row", function () {
        const rowIndex = parseInt($(this).data("row-index"));
        const empName = $(this).data("emp-name") || "";
        const empCode = $(this).data("emp-code") || "";

        if (!confirm(`Are you sure you want to delete ${empName} (${empCode}) from the headcount?\n\nThis action cannot be undone.`)) return;

        const $btn = $(this);
        $btn.prop("disabled", true).html('<i class="bi bi-hourglass-split"></i>');

        $.ajax({
            url: `/api/headcount/${rowIndex}`,
            method: "DELETE",
            success: function (resp) {
                toastr.success(resp.message || "Employee deleted successfully");
                loadData();
                loadSummary();
                loadDemandSummary();
                if ($("#pane-demand").hasClass("show")) loadDemandData();
            },
            error: function (xhr) {
                const msg = xhr.responseJSON ? xhr.responseJSON.error : "Delete failed";
                toastr.error(msg);
                $btn.prop("disabled", false).html('<i class="bi bi-trash"></i>');
            }
        });
    });

    $("#btnSaveEdit").click(saveEdit);

    $(document).on("click", "#tableHeader th[data-col]", function () {
        const col = $(this).data("col");
        if (sortColumn === col) { sortAsc = !sortAsc; }
        else { sortColumn = col; sortAsc = true; }
        currentData.sort(function (a, b) {
            const va = a[col] != null ? String(a[col]) : "";
            const vb = b[col] != null ? String(b[col]) : "";
            return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
        });
        renderTableHeader();
        renderTableBody(currentData);
    });

    $("#btnRefresh").click(function () {
        const $btn = $(this);
        $btn.prop("disabled", true).html('<i class="bi bi-hourglass-split me-1"></i>Refreshing...');
        $.post("/api/refresh", function (resp) {
            toastr.success(`Refreshed: ${resp.total_rows} rows loaded`);
            currentPage = 1;
            demandPage = 1;
            loadData();
            loadSummary();
            loadDemandSummary();
            if ($("#pane-demand").hasClass("show")) loadDemandData();
        }).fail(function () {
            toastr.error("Refresh failed");
        }).always(function () {
            $btn.prop("disabled", false).html('<i class="bi bi-arrow-clockwise me-1"></i>Refresh');
        });
    });

    $("#btnExport").click(exportCSV);

    // ── Demand Fulfillment Tab Events ─────────────────────────
    $('button[data-bs-toggle="tab"]').on("shown.bs.tab", function (e) {
        if (e.target.id === "tab-demand") {
            loadDemandData();
        }
    });

    $("#btnDemandApply").click(function () { demandPage = 1; loadDemandData(); });

    $("#btnDemandReset").click(function () {
        $("#filterDemandStatus").val("All");
        $("#demandSearchInput").val("");
        activeDfKpi = "all";
        $(".df-card-clickable").removeClass("active");
        $('[data-df-kpi="all"]').addClass("active");
        demandPage = 1;
        loadDemandData();
    });

    $("#demandSearchInput").on("keydown", function (e) {
        if (e.key === "Enter") { demandPage = 1; loadDemandData(); }
    });

    $("#demandPerPageSelect").change(function () {
        demandPerPage = parseInt($(this).val());
        demandPage = 1;
        loadDemandData();
    });

    $(document).on("click", "#demandPaginationControls .page-link", function (e) {
        e.preventDefault();
        const p = parseInt($(this).data("page"));
        if (!isNaN(p) && p >= 1) { demandPage = p; loadDemandData(); }
    });

    $(document).on("click", ".btn-find-matches", function () {
        openSuggestionsModal($(this).data("row"));
    });

    $(document).on("click", ".btn-map-employee", function () {
        mapEmployee(
            parseInt($(this).data("emp-row")),
            $(this).data("emp-name"),
            $(this).data("emp-code")
        );
    });

    $(document).on("click", ".btn-mark-external", function () {
        markAsExternal(parseInt($(this).data("row-index")));
    });

    // ── Add External Employee ────────────────────────────────
    $(document).on("click", ".btn-add-external", function () {
        const row = $(this).data("row");
        $("#extDemandRowIndex").val(row.row_index);
        $("#extEmpReqId").text("— " + (row["Requisition ID"] || ""));
        $("#extEmpCode").val("");
        $("#extEmpName").val("");
        $("#extSkills").val(row["Skillset"] || "");
        $("#extSubPractice").val("");
        $("#extExperience").val("");
        $("#extDesignation").val("");
        $("#extCustomerName").val(row["Customer Name"] || "");
        $("#extProjects").val("");
        new bootstrap.Modal(document.getElementById("externalEmpModal")).show();
    });

    $("#btnSaveExternalEmp").click(function () {
        const empCode = $("#extEmpCode").val().trim();
        const empName = $("#extEmpName").val().trim();
        const skills = $("#extSkills").val().trim();
        const subPractice = $("#extSubPractice").val();
        const demandRowIndex = parseInt($("#extDemandRowIndex").val());

        if (!empCode || !empName || !subPractice) {
            toastr.warning("Please fill in all required fields (Emp Code, Emp Name, Sub Practice)");
            return;
        }

        if (!confirm(`Add external employee ${empName} (${empCode}) to Head Count and mark demand as Fulfilled?`)) return;

        showLoading();
        $.ajax({
            url: "/api/demand-fulfillment/fulfill-external",
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify({
                demand_row_index: demandRowIndex,
                emp_code: empCode,
                emp_name: empName,
                skills: skills,
                sub_practice: subPractice,
                experience: $("#extExperience").val().trim(),
                designation: $("#extDesignation").val().trim(),
                customer_name: $("#extCustomerName").val().trim(),
                projects: $("#extProjects").val().trim(),
            }),
            success: function (resp) {
                hideLoading();
                bootstrap.Modal.getInstance(document.getElementById("externalEmpModal")).hide();
                toastr.success(resp.message || "External employee added and demand fulfilled!");
                loadDemandData();
                loadDemandSummary();
                loadData();
                loadSummary();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Failed to add external employee";
                toastr.error(err);
            }
        });
    });

    $(document).on("click", ".btn-confirm-demand", function () {
        const rowIndex = parseInt($(this).data("row-index"));
        if (!confirm("Confirm this fulfillment?\n\nThe employee will be marked as Billable and the demand as Fulfilled.")) return;

        const $btn = $(this);
        $btn.prop("disabled", true).html('<i class="bi bi-hourglass-split me-1"></i>Confirming...');

        $.ajax({
            url: "/api/demand-fulfillment/confirm",
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify({ demand_row_index: rowIndex }),
            success: function (resp) {
                toastr.success(resp.message || "Demand fulfilled successfully");
                loadDemandData();
                loadDemandSummary();
                loadData();
                loadSummary();
            },
            error: function (xhr) {
                const msg = xhr.responseJSON ? xhr.responseJSON.error : "Confirm failed";
                toastr.error(msg);
                $btn.prop("disabled", false).html('<i class="bi bi-check2-circle me-1"></i>Confirm');
            }
        });
    });

    $(document).on("click", ".btn-delete-demand", function () {
        const rowIndex = parseInt($(this).data("row-index"));
        const reqId = $(this).data("req-id") || rowIndex;
        const isMapped = $(this).data("is-mapped") === "1" || $(this).data("is-mapped") === 1;

        let confirmMsg = `Are you sure you want to delete demand "${reqId}"?`;
        if (isMapped) {
            confirmMsg += "\n\nThis demand has a mapped employee. Their status will be reverted to Non-Billable.";
        }

        if (!confirm(confirmMsg)) return;

        const $btn = $(this);
        $btn.prop("disabled", true).html('<i class="bi bi-hourglass-split"></i>');

        $.ajax({
            url: `/api/demand-requisition/${rowIndex}`,
            method: "DELETE",
            success: function (resp) {
                toastr.success(resp.message || "Demand deleted successfully");
                loadDemandData();
                loadDemandSummary();
                if (isMapped) {
                    loadData();
                }
            },
            error: function (xhr) {
                const msg = xhr.responseJSON ? xhr.responseJSON.error : "Delete failed";
                toastr.error(msg);
                $btn.prop("disabled", false).html('<i class="bi bi-trash"></i>');
            }
        });
    });

    // ── Demand Requisition Form ───────────────────────────────
    $("#demandRequisitionForm").on("submit", function (e) {
        e.preventDefault();

        const requisitionId = $("#drRequisitionId").val().trim();
        const skillset = $("#drSkillset").val().trim();
        const demandStatus = $("#drDemandStatus").val();
        const customerName = $("#drCustomerName").val().trim();

        if (!requisitionId || !skillset || !demandStatus || !customerName) {
            toastr.warning("Please fill in all required fields");
            return;
        }

        const minExp = $("#drMinExp").val();
        const maxExp = $("#drMaxExp").val();
        let yrsOfExp = "";
        if (minExp && maxExp) yrsOfExp = `${minExp}-${maxExp}`;
        else if (minExp) yrsOfExp = `${minExp}+`;
        else if (maxExp) yrsOfExp = `0-${maxExp}`;

        const payload = {
            "Requisition ID": requisitionId,
            "Yrs of Exp": yrsOfExp,
            "Skillset": skillset,
            "Demand Status": demandStatus,
            "Notes": $("#drNotes").val().trim(),
            "Customer Name": $("#drCustomerName").val().trim(),
        };

        showLoading();
        $.ajax({
            url: "/api/demand-requisition",
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify(payload),
            success: function (resp) {
                hideLoading();
                toastr.success(resp.message || "Saved to SharePoint!");
                $("#demandRequisitionForm")[0].reset();
                loadDemandSummary();
                if ($("#pane-demand").hasClass("show")) loadDemandData();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Save failed";
                toastr.error(err);
            }
        });
    });

    // ── Add Employees (Bulk) ──────────────────────────────────
    const NEW_EMP_FIELDS = [
        "Division", "Emp Code", "Emp Name", "Status", "LWD", "Skills",
        "Fresher/Lateral", "Offshore/Onsite", "Experience", "Designation",
        "Grade", "DOJ", "Gender", "First Line Manager", "Skip Level Manager",
        "Company Email", "Sub Practice", "Remarks", "Empower SL",
        "Billable/Non Billable", "Billable Till Date", "Projects", "Remarks2",
        "Customer Name", "Customer interview happened(Yes/No)",
        "Customer Selected(Yes/No)", "Comments",
    ];

    const NEW_EMP_PLACEHOLDERS = {
        "Division": "",
        "Emp Code": "e.g. 20001",
        "Emp Name": "Full name",
        "Status": "e.g. Confirmed",
        "LWD": "YYYY-MM-DD",
        "Skills": "Python, SQL",
        "Fresher/Lateral": "e.g. Lateral",
        "Offshore/Onsite": "e.g. Offshore",
        "Experience": "e.g. 5:3:0",
        "Designation": "e.g. Senior Consultant",
        "Grade": "",
        "DOJ": "YYYY-MM-DD",
        "Gender": "",
        "First Line Manager": "",
        "Skip Level Manager": "",
        "Company Email": "name@rsystems.com",
        "Sub Practice": "e.g. AI / BI / DE",
        "Remarks": "",
        "Empower SL": "",
        "Billable/Non Billable": "e.g. Non-Billable",
        "Billable Till Date": "YYYY-MM-DD",
        "Projects": "",
        "Remarks2": "",
        "Customer Name": "",
        "Customer interview happened(Yes/No)": "Yes / No",
        "Customer Selected(Yes/No)": "Yes / No",
        "Comments": "",
    };

    function newEmpRowHtml(prefill) {
        const p = prefill || {};
        const v = (k) => escapeHtml(p[k] || "");
        const cells = NEW_EMP_FIELDS.map(field => {
            const ph = NEW_EMP_PLACEHOLDERS[field] || "";
            return `<td><input class="form-control form-control-sm" `
                + `data-field="${field}" value="${v(field)}" `
                + `placeholder="${escapeHtml(ph)}"></td>`;
        }).join("");
        return `
            <tr class="new-emp-row">
                <td class="text-center">
                    <button type="button" class="btn btn-link btn-sm text-danger p-0 btn-remove-new-emp"
                            title="Remove row">
                        <i class="bi bi-x-circle-fill"></i>
                    </button>
                </td>
                ${cells}
            </tr>`;
    }

    function addNewEmpRow(prefill) {
        $("#newEmpTableBody").append(newEmpRowHtml(prefill));
        updateNewEmpRowCount();
    }

    function updateNewEmpRowCount() {
        $("#newEmpRowCount").text($("#newEmpTableBody .new-emp-row").length);
    }

    function resetAddEmployeesModal() {
        $("#newEmpTableBody").empty();
        $("#bulkPasteArea").val("");
        $("#bulkPasteWrap").addClass("d-none");
        addNewEmpRow();
    }

    $("#addEmployeesModal").on("show.bs.modal", function () {
        if ($("#newEmpTableBody .new-emp-row").length === 0) {
            addNewEmpRow();
        }
    });

    $("#addEmployeesModal").on("hidden.bs.modal", function () {
        resetAddEmployeesModal();
        $("#newEmpTableBody").empty();
        updateNewEmpRowCount();
    });

    $("#btnAddEmpRow").on("click", function () {
        addNewEmpRow();
    });

    $(document).on("click", ".btn-remove-new-emp", function () {
        $(this).closest("tr").remove();
        updateNewEmpRowCount();
        if ($("#newEmpTableBody .new-emp-row").length === 0) {
            addNewEmpRow();
        }
    });

    $("#btnClearNewEmp").on("click", function () {
        if (!confirm("Clear all rows?")) return;
        $("#newEmpTableBody").empty();
        addNewEmpRow();
    });

    $("#btnBulkPasteToggle").on("click", function () {
        $("#bulkPasteWrap").toggleClass("d-none");
        if (!$("#bulkPasteWrap").hasClass("d-none")) {
            $("#bulkPasteArea").trigger("focus");
        }
    });

    $("#btnBulkPasteCancel").on("click", function () {
        $("#bulkPasteWrap").addClass("d-none");
        $("#bulkPasteArea").val("");
    });

    $("#btnBulkPasteApply").on("click", function () {
        const raw = $("#bulkPasteArea").val();
        if (!raw || !raw.trim()) {
            toastr.warning("Nothing to paste");
            return;
        }
        const headers = NEW_EMP_FIELDS;
        const lines = raw.split(/\r?\n/).filter(l => l.trim().length > 0);

        const $body = $("#newEmpTableBody");
        $body.find(".new-emp-row").each(function () {
            const hasData = $(this).find("[data-field]").toArray()
                .some(el => $(el).val() && $(el).val().trim());
            if (!hasData) $(this).remove();
        });

        let added = 0;
        lines.forEach(line => {
            const cells = line.split("\t");
            const prefill = {};
            headers.forEach((h, i) => {
                if (cells[i] !== undefined) prefill[h] = cells[i].trim();
            });
            addNewEmpRow(prefill);
            added++;
        });

        $("#bulkPasteArea").val("");
        $("#bulkPasteWrap").addClass("d-none");
        toastr.success(`Added ${added} row(s) from paste`);
    });

    function collectNewEmployees() {
        const rows = [];
        $("#newEmpTableBody .new-emp-row").each(function () {
            const obj = {};
            let anyValue = false;
            $(this).find("[data-field]").each(function () {
                const field = $(this).data("field");
                const val = ($(this).val() || "").trim();
                obj[field] = val;
                if (val) anyValue = true;
            });
            if (anyValue) rows.push(obj);
        });
        return rows;
    }

    $("#btnSaveNewEmployees").on("click", function () {
        const employees = collectNewEmployees();
        if (employees.length === 0) {
            toastr.warning("Please fill in at least one employee");
            return;
        }

        const required = ["Emp Code", "Emp Name", "Sub Practice"];
        const seenCodes = new Set();
        for (let i = 0; i < employees.length; i++) {
            const e = employees[i];
            const missing = required.filter(f => !e[f]);
            if (missing.length) {
                toastr.error(`Row ${i + 1}: missing ${missing.join(", ")}`);
                return;
            }
            const code = e["Emp Code"].trim();
            if (seenCodes.has(code)) {
                toastr.error(`Row ${i + 1}: duplicate Emp Code "${code}" in this batch`);
                return;
            }
            seenCodes.add(code);
        }

        if (!confirm(`Save ${employees.length} employee(s) to Head Count Report on SharePoint?`)) return;

        showLoading();
        $.ajax({
            url: "/api/headcount/bulk-add",
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify({ employees: employees }),
            success: function (resp) {
                hideLoading();
                bootstrap.Modal.getInstance(document.getElementById("addEmployeesModal")).hide();
                toastr.success(resp.message || `${resp.added} employee(s) added`);
                loadData();
                loadSummary();
                loadFilters();
            },
            error: function (xhr) {
                hideLoading();
                const err = (xhr.responseJSON && xhr.responseJSON.error) || "Failed to add employees";
                toastr.error(err);
            }
        });
    });

    // ── Manager Notifications ─────────────────────────────────
    function openNotifyModal(empRow, empName, empCode) {
        $("#notifyEmpLabel").text(`${empName} (${empCode})`);
        $("#notifyEmpRowIndex").val(empRow);
        $("#notifyTo").val("");
        $("#notifyCc").val("");
        $("#notifySubject").val("");
        $("#notifyBody").val("");
        $("#notifyPreview").html('<em class="text-muted">Loading…</em>');
        $("#notifyResolutionAlert").addClass("d-none").removeClass("alert-success alert-warning alert-danger alert-info");
        $("#notifyDemandLinkAlert").addClass("d-none").empty();

        const modal = new bootstrap.Modal(document.getElementById("notifyManagerModal"));
        modal.show();

        $.getJSON(`/api/notify-manager/preview/${empRow}`, function (resp) {
            const r = resp.manager_resolution || {};
            $("#notifyManagerName").val(r.manager_name || "");
            $("#notifyManagerNameLabel").text(
                r.manager_name ? `Manager: ${r.manager_name}` : "Manager not specified on row"
            );
            $("#notifyResolutionMethod").val(r.method || "manual");
            $("#notifyTo").val(resp.default_to || "");
            $("#notifyCc").val((resp.default_cc || []).join(", "));
            $("#notifySubject").val(resp.subject || "");
            $("#notifyBody").val(resp.body_html || "");
            renderNotifyPreview();

            const $dlAlert = $("#notifyDemandLinkAlert");
            if (resp.has_demand_link) {
                $dlAlert.addClass("d-none").empty();
            } else {
                $dlAlert.removeClass("d-none").html(
                    '<i class="bi bi-info-circle-fill me-1"></i>'
                    + 'This employee is <b>Proposed</b> but not linked to any '
                    + 'open Demand Requisition. The email will be sent without '
                    + 'requisition details. On Approve, the employee will be '
                    + 'marked Billable directly (no demand row to update).'
                );
            }

            const $a = $("#notifyResolutionAlert").removeClass("d-none alert-success alert-warning alert-danger alert-info");
            const status = r.status;
            if (status === "ok") {
                $a.addClass("alert-success").html(
                    `<i class="bi bi-check-circle-fill me-1"></i>Manager email auto-resolved (exact match).`
                );
            } else if (status === "ok_disambiguated") {
                $a.addClass("alert-success").html(
                    `<i class="bi bi-check-circle-fill me-1"></i>Manager email auto-resolved by Sub Practice match.`
                );
            } else if (status === "ok_fuzzy") {
                $a.addClass("alert-warning").html(
                    `<i class="bi bi-exclamation-triangle-fill me-1"></i>Fuzzy match (score ${r.fuzzy_score}). Please verify the email before sending.`
                );
            } else if (status === "multiple_matches") {
                const opts = (r.candidates || []).map(c => `${c[0]} <${c[1] || 'no email'}>`).join("; ");
                $a.addClass("alert-warning").html(
                    `<i class="bi bi-exclamation-triangle-fill me-1"></i>Multiple employees share this manager name. Please pick the correct email.<br><small>Candidates: ${escapeHtml(opts)}</small>`
                );
            } else if (status === "no_email") {
                $a.addClass("alert-warning").html(
                    `<i class="bi bi-exclamation-triangle-fill me-1"></i>Matched a manager row but their <code>Company Email</code> is blank. Please enter manually.`
                );
            } else if (status === "no_match" || status === "not_found") {
                $a.addClass("alert-danger").html(
                    `<i class="bi bi-x-circle-fill me-1"></i>Could not auto-resolve manager email — please enter it manually.`
                );
            }
        }).fail(function (xhr) {
            const err = xhr.responseJSON ? xhr.responseJSON.error : "Failed to load preview";
            toastr.error(err);
            modal.hide();
        });
    }

    function renderNotifyPreview() {
        $("#notifyPreview").html($("#notifyBody").val() || '<em class="text-muted">(empty)</em>');
    }

    function sendNotifyMail() {
        const empRow = parseInt($("#notifyEmpRowIndex").val());
        const to = ($("#notifyTo").val() || "").trim();
        const cc = ($("#notifyCc").val() || "").split(",").map(s => s.trim()).filter(Boolean);
        const subject = ($("#notifySubject").val() || "").trim();
        const body = $("#notifyBody").val() || "";

        if (!to) { toastr.warning("Manager email is required"); return; }
        if (!subject) { toastr.warning("Subject is required"); return; }
        if (!body.trim()) { toastr.warning("Body cannot be empty"); return; }
        if (!confirm(`Send approval email to ${to}?`)) return;

        showLoading();
        $.ajax({
            url: "/api/notify-manager",
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify({
                emp_row_index: empRow,
                manager_email: to,
                manager_name: $("#notifyManagerName").val() || "",
                cc_emails: cc,
                subject: subject,
                body_html: body,
                resolution_method: $("#notifyResolutionMethod").val() || "manual",
            }),
            success: function (resp) {
                hideLoading();
                bootstrap.Modal.getInstance(document.getElementById("notifyManagerModal")).hide();
                toastr.success(resp.message || "Email sent");
                loadData();
                refreshNotifAwaitingBadge();
                if ($("#pane-notifications").hasClass("show")) loadNotifications();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Send failed";
                toastr.error(err);
            }
        });
    }

    // ── Notifications Tab ─────────────────────────────────────
    function loadNotifications() {
        const status = $("#notifFilterStatus").val() || "";
        const url = "/api/notifications" + (status ? `?status=${encodeURIComponent(status)}` : "");
        $.getJSON(url, function (resp) {
            renderNotifTable(resp.data || []);
            $("#notifTotalLabel").text(`${resp.total} notification(s)`);
        }).fail(function () {
            toastr.error("Failed to load notifications");
        });
    }

    function renderNotifTable(rows) {
        if (!rows.length) {
            $("#notifTableBody").html(
                '<tr><td colspan="10" class="text-center py-3 text-muted small">No notifications</td></tr>'
            );
            return;
        }
        let html = "";
        rows.forEach(function (n, i) {
            const empLbl = `${escapeHtml(n.emp_name || "")} <span class="text-muted">(${escapeHtml(n.emp_code || "")})</span>`;
            const sent = n.sent_at ? n.sent_at.replace("T", " ").substr(0, 16) : "";
            const typeBadge = n.notification_type === "blocked"
                ? '<span class="badge bg-warning text-dark">Blocked</span>'
                : '<span class="badge bg-info text-dark">Proposed</span>';
            html += `<tr>
                <td class="small text-muted">${i + 1}</td>
                <td class="small">${empLbl}</td>
                <td class="small text-center">${typeBadge}</td>
                <td class="small">${escapeHtml(n.demand_req_id || "—")}</td>
                <td class="small">${escapeHtml(n.customer_name || "—")}</td>
                <td class="small">${escapeHtml(n.manager_name || "")}<br><span class="text-muted">${escapeHtml(n.manager_email || "")}</span></td>
                <td class="small">${escapeHtml(sent)}</td>
                <td class="small text-center">${n.reminder_count || 0}</td>
                <td class="small">${renderNotifBadge(n.status)}</td>
                <td class="text-center text-nowrap">
                    ${notifRowActions(n)}
                </td>
            </tr>`;
        });
        $("#notifTableBody").html(html);
    }

    function notifRowActions(n) {
        const deleteBtn = `<button class="btn btn-outline-danger btn-sm btn-notif-delete" data-id="${n.id}" title="Delete notification">
                <i class="bi bi-trash"></i>
            </button>`;

        if (n.status === "awaiting_reply") {
            const isBlocked = n.notification_type === "blocked";
            if (isBlocked) {
                return `
                    <button class="btn btn-success btn-sm btn-notif-blocked-approve" data-id="${n.id}" title="Allocation completed">
                        <i class="bi bi-check-circle"></i>
                    </button>
                    <button class="btn btn-secondary btn-sm btn-notif-blocked-reject" data-id="${n.id}" title="Not allocated">
                        <i class="bi bi-x-circle"></i>
                    </button>
                    <button class="btn btn-outline-warning btn-sm btn-notif-resend" data-id="${n.id}" title="Resend now">
                        <i class="bi bi-arrow-repeat"></i>
                    </button>
                    <button class="btn btn-outline-secondary btn-sm btn-notif-cancel" data-id="${n.id}" title="Stop reminders">
                        <i class="bi bi-bell-slash"></i>
                    </button>
                    ${deleteBtn}`;
            }
            return `
                <button class="btn btn-success btn-sm btn-notif-approve" data-id="${n.id}" title="Manager approved">
                    <i class="bi bi-check-circle"></i>
                </button>
                <button class="btn btn-danger btn-sm btn-notif-reject" data-id="${n.id}" title="Manager rejected">
                    <i class="bi bi-x-circle"></i>
                </button>
                <button class="btn btn-outline-warning btn-sm btn-notif-resend" data-id="${n.id}" title="Resend now">
                    <i class="bi bi-arrow-repeat"></i>
                </button>
                <button class="btn btn-outline-secondary btn-sm btn-notif-cancel" data-id="${n.id}" title="Stop reminders">
                    <i class="bi bi-bell-slash"></i>
                </button>
                ${deleteBtn}`;
        }

        let overrideBtns = "";
        if (n.status === "approved") {
            overrideBtns = `<button class="btn btn-outline-danger btn-sm btn-notif-override" data-id="${n.id}" data-action="reject" title="Override → Reject (Non-Billable)">
                    <i class="bi bi-arrow-counterclockwise me-1"></i>Reject
                </button>`;
        } else if (n.status === "rejected" || n.status === "no_response") {
            overrideBtns = `<button class="btn btn-outline-success btn-sm btn-notif-override" data-id="${n.id}" data-action="approve" title="Override → Approve (Billable)">
                    <i class="bi bi-arrow-counterclockwise me-1"></i>Approve
                </button>`;
        }

        return `${overrideBtns} ${deleteBtn}`;
    }

    function refreshNotifAwaitingBadge() {
        $.getJSON("/api/notifications?status=awaiting_reply", function (resp) {
            const $b = $("#notifAwaitingBadge");
            if ((resp.total || 0) > 0) {
                $b.text(resp.total).removeClass("d-none");
            } else {
                $b.text("").addClass("d-none");
            }
        });
    }

    // ── Blocked Allocation Check ────────────────────────────────
    function openBlockedModal(empRow, empName, empCode) {
        $("#blockedEmpLabel").text(`${empName} (${empCode})`);
        $("#blockedEmpRowIndex").val(empRow);
        $("#blockedTo").val("");
        $("#blockedCc").val("");
        $("#blockedSubject").val("");
        $("#blockedBody").val("");
        $("#blockedPreview").html('<em class="text-muted">Loading…</em>');
        $("#blockedResolutionAlert").addClass("d-none").removeClass("alert-success alert-warning alert-danger alert-info");

        const modal = new bootstrap.Modal(document.getElementById("blockedAllocationModal"));
        modal.show();

        $.getJSON(`/api/notify-blocked/preview/${empRow}`, function (resp) {
            const r = resp.manager_resolution || {};
            $("#blockedManagerName").val(r.manager_name || "");
            $("#blockedManagerNameLabel").text(
                r.manager_name ? `Manager: ${r.manager_name}` : "Manager not specified on row"
            );
            $("#blockedResolutionMethod").val(r.method || "manual");
            $("#blockedTo").val(resp.default_to || "");
            $("#blockedCc").val((resp.default_cc || []).join(", "));
            $("#blockedSubject").val(resp.subject || "");
            $("#blockedBody").val(resp.body_html || "");
            renderBlockedPreview();

            const $a = $("#blockedResolutionAlert").removeClass("d-none alert-success alert-warning alert-danger alert-info");
            const status = r.status;
            if (status === "ok" || status === "ok_disambiguated") {
                $a.addClass("alert-success").html(
                    `<i class="bi bi-check-circle-fill me-1"></i>Manager email auto-resolved.`
                );
            } else if (status === "ok_fuzzy") {
                $a.addClass("alert-warning").html(
                    `<i class="bi bi-exclamation-triangle-fill me-1"></i>Fuzzy match (score ${r.fuzzy_score}). Please verify the email.`
                );
            } else if (status === "multiple_matches") {
                const opts = (r.candidates || []).map(c => `${c[0]} <${c[1] || 'no email'}>`).join("; ");
                $a.addClass("alert-warning").html(
                    `<i class="bi bi-exclamation-triangle-fill me-1"></i>Multiple matches. Please pick the correct email.<br><small>Candidates: ${escapeHtml(opts)}</small>`
                );
            } else {
                $a.addClass("alert-danger").html(
                    `<i class="bi bi-x-circle-fill me-1"></i>Could not auto-resolve manager email — please enter it manually.`
                );
            }
        }).fail(function (xhr) {
            const err = xhr.responseJSON ? xhr.responseJSON.error : "Failed to load preview";
            toastr.error(err);
            modal.hide();
        });
    }

    function renderBlockedPreview() {
        $("#blockedPreview").html($("#blockedBody").val() || '<em class="text-muted">(empty)</em>');
    }

    function sendBlockedMail() {
        const empRow = parseInt($("#blockedEmpRowIndex").val());
        const to = ($("#blockedTo").val() || "").trim();
        const cc = ($("#blockedCc").val() || "").split(",").map(s => s.trim()).filter(Boolean);
        const subject = ($("#blockedSubject").val() || "").trim();
        const body = $("#blockedBody").val() || "";

        if (!to) { toastr.warning("Manager email is required"); return; }
        if (!subject) { toastr.warning("Subject is required"); return; }
        if (!body.trim()) { toastr.warning("Body cannot be empty"); return; }
        if (!confirm(`Send allocation check email to ${to}?`)) return;

        showLoading();
        $.ajax({
            url: "/api/notify-blocked",
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify({
                emp_row_index: empRow,
                manager_email: to,
                manager_name: $("#blockedManagerName").val() || "",
                cc_emails: cc,
                subject: subject,
                body_html: body,
                resolution_method: $("#blockedResolutionMethod").val() || "manual",
            }),
            success: function (resp) {
                hideLoading();
                bootstrap.Modal.getInstance(document.getElementById("blockedAllocationModal")).hide();
                toastr.success(resp.message || "Email sent");
                loadData();
                refreshNotifAwaitingBadge();
                if ($("#pane-notifications").hasClass("show")) loadNotifications();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Send failed";
                toastr.error(err);
            }
        });
    }

    $(document).on("click", ".btn-check-allocation", function () {
        const empRow = parseInt($(this).data("emp-row"));
        const empName = $(this).data("emp-name") || "";
        const empCode = $(this).data("emp-code") || "";
        openBlockedModal(empRow, empName, empCode);
    });

    $("#btnBlockedPreview").click(renderBlockedPreview);
    $("#blockedBody").on("input blur", renderBlockedPreview);
    $("#btnBlockedSend").click(sendBlockedMail);

    $("#blockedAllocationModal").on("show.bs.modal", function () {
        $("#blockedReminderDays").text("7");
        $("#blockedMaxReminders").text("3");
    });

    // ── Notification action handlers ──────────────────────────
    $(document).on("click", ".btn-notify-manager", function () {
        const empRow = parseInt($(this).data("emp-row"));
        const empName = $(this).data("emp-name") || "";
        const empCode = $(this).data("emp-code") || "";
        openNotifyModal(empRow, empName, empCode);
    });

    $("#btnNotifyPreview").click(renderNotifyPreview);
    $("#notifyBody").on("input blur", renderNotifyPreview);
    $("#btnNotifySend").click(sendNotifyMail);

    $("#notifyManagerModal").on("show.bs.modal", function () {
        $("#notifyReminderDays").text("7");
        $("#notifyMaxReminders").text("3");
    });

    $("#btnNotifRefresh").click(loadNotifications);
    $("#notifFilterStatus").change(loadNotifications);

    $('button[data-bs-toggle="tab"]').on("shown.bs.tab", function (e) {
        if (e.target.id === "tab-notifications") {
            loadNotifications();
            refreshNotifAwaitingBadge();
        }
    });

    $(document).on("click", ".btn-notif-approve", function () {
        const id = $(this).data("id");
        const note = prompt("Optional approval note:", "");
        if (note === null) return;
        if (!confirm("Mark as Approved? This will set the employee to Billable and the demand to Fulfilled.")) return;
        notifAction(id, "approve", { note: note });
    });

    $(document).on("click", ".btn-notif-reject", function () {
        const id = $(this).data("id");
        const note = prompt("Optional rejection note:", "");
        if (note === null) return;
        if (!confirm("Mark as Rejected? This will revert the employee to Non-Billable and reopen the demand.")) return;
        notifAction(id, "reject", { note: note });
    });

    $(document).on("click", ".btn-notif-cancel", function () {
        const id = $(this).data("id");
        if (!confirm("Stop sending reminders for this notification? No data will be changed.")) return;
        notifAction(id, "cancel", {});
    });

    $(document).on("click", ".btn-notif-blocked-approve", function () {
        const id = $(this).data("id");
        const allocDate = prompt("Enter allocation completion date (YYYY-MM-DD):", new Date().toISOString().slice(0, 10));
        if (allocDate === null) return;
        if (!allocDate.trim()) { toastr.warning("Allocation date is required"); return; }
        if (!confirm("Mark allocation as completed? Employee will be set to Billable.")) return;
        notifAction(id, "blocked-approve", { allocation_date: allocDate.trim() });
    });

    $(document).on("click", ".btn-notif-blocked-reject", function () {
        const id = $(this).data("id");
        const note = prompt("Optional note:", "");
        if (note === null) return;
        if (!confirm("Mark as not allocated? Employee will remain Blocked.")) return;
        notifAction(id, "blocked-reject", { note: note });
    });

    $(document).on("click", ".btn-notif-delete", function () {
        const id = $(this).data("id");
        if (!confirm("Are you sure you want to permanently delete this notification? This cannot be undone.")) return;
        notifAction(id, "delete", {});
    });

    $(document).on("click", ".btn-notif-override", function () {
        const id = $(this).data("id");
        const action = $(this).data("action");
        const label = action === "approve" ? "Approve (Billable)" : "Reject (Non-Billable)";
        if (!confirm(`Override this notification to ${label}? This will update the employee's status in SharePoint.`)) return;
        const note = prompt("Optional note for the override:", "");
        if (note === null) return;
        showLoading();
        $.ajax({
            url: `/api/notifications/${id}/override`,
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify({ action: action, note: note }),
            success: function (resp) {
                hideLoading();
                toastr.success(resp.message || "Override applied");
                loadNotifications();
                refreshNotifAwaitingBadge();
                loadData();
                loadSummary();
                if ($("#pane-demand").hasClass("show")) loadDemandData();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Override failed";
                toastr.error(err);
            }
        });
    });

    $(document).on("click", ".btn-notif-resend", function () {
        const id = $(this).data("id");
        if (!confirm("Resend the original email and reset the reminder counter?")) return;
        notifAction(id, "resend", {});
    });

    function notifAction(id, action, body) {
        showLoading();
        $.ajax({
            url: `/api/notifications/${id}/${action}`,
            method: "POST",
            contentType: "application/json",
            data: JSON.stringify(body || {}),
            success: function (resp) {
                hideLoading();
                toastr.success(resp.message || "Done");
                loadNotifications();
                refreshNotifAwaitingBadge();
                loadData();
                loadSummary();
                if ($("#pane-demand").hasClass("show")) loadDemandData();
            },
            error: function (xhr) {
                hideLoading();
                const err = xhr.responseJSON ? xhr.responseJSON.error : "Action failed";
                toastr.error(err);
            }
        });
    }

    // ── Helpers ───────────────────────────────────────────────
    function getFilterParams() {
        return {
            sub_practice: $("#filterSubPractice").val() || "All",
            billable: $("#filterBillable").val() || "All",
            project: $("#filterProject").val() || "All",
            search: $("#searchInput").val() || "",
            kpi: activeKpi,
        };
    }

    function escapeHtml(str) {
        const div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    }

    function showLoading() {
        $("#loadingOverlay").removeClass("d-none").addClass("d-flex");
    }

    function hideLoading() {
        $("#loadingOverlay").removeClass("d-flex").addClass("d-none");
    }

    // Rotate chevron icons on KPI section collapse/expand
    $(".kpi-section-header").on("click", function () {
        var $icon = $(this).find(".kpi-collapse-icon");
        var target = $(this).data("bs-target");
        var $target = $(target);
        $target.on("shown.bs.collapse hidden.bs.collapse", function () {
            if ($target.hasClass("show")) {
                $icon.css("transform", "rotate(0deg)");
            } else {
                $icon.css("transform", "rotate(180deg)");
            }
            $target.off("shown.bs.collapse hidden.bs.collapse");
        });
    });
});
