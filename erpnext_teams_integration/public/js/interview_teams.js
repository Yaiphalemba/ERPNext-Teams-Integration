let teams_settings_cache = null;

async function get_teams_settings() {
    if (!teams_settings_cache) {
        teams_settings_cache = await frappe.db.get_doc(
            "Teams Settings",
            "Teams Settings"
        );
    }
    return teams_settings_cache;
}

frappe.ui.form.on("Interview", {
    async refresh(frm) {
        const settings_doc = await get_teams_settings();

        const enabledSet = new Set(
            (settings_doc.enabled_doctypes || [])
                .map(row => row.doctype_name)
        );

        if (!enabledSet.has(frm.doctype)) {
            frm.remove_custom_button(__('Create Teams Meeting'), __("Teams"));
            frm.remove_custom_button(__('Reschedule Teams Meeting'), __("Teams"));
            frm.remove_custom_button(__('Cancel Teams Meeting'), __("Teams"));
            return;
        }

        if (frm.doc.custom_teams_meeting_url) {
            frm.fields_dict.custom_join_teams_meeting.$wrapper
                .find("button")
                .off("click") // Remove existing handler if any
                .on("click", function () {
                    window.open(frm.doc.custom_teams_meeting_url, "_blank");
                });
        }
        
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get("teams_authentication_status") === "success") {
            frappe.msgprint({
                title: "Token Fetched Successfully",
                message: "Teams token was successfully saved after login.",
                indicator: 'green'
            });
            const cleanURL = new URL(window.location.href);
            cleanURL.searchParams.delete('teams_authentication_status');
            window.history.replaceState({}, document.title, cleanURL.pathname);
        }

        if (!frm.doc.__islocal) {
            // Create a dropdown called "Teams"
            frm.add_custom_button(__('Create Teams Meeting'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.create_meeting",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        if (r.message) {
                            // If it's an object, show the .message field
                            let msg = (typeof r.message === "string") ? r.message : r.message.message;
                            frappe.msgprint(msg);
                        } else if (r.message && r.message.login_url) {
                            // Redirect to MS login if required
                            window.location.href = r.message.login_url;
                        }
                        frm.reload_doc();
                    }
                });
            }, __("Teams"));

            frm.add_custom_button(__('Reschedule Teams Meeting'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.reschedule_meeting",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        if (r.message) {
                            // If it's an object, show the .message field
                            let msg = (typeof r.message === "string") ? r.message : r.message.message;
                            frappe.msgprint(msg);
                        } else if (r.message && r.message.login_url) {
                            // Redirect to MS login if required
                            window.location.href = r.message.login_url;
                        }
                        frm.reload_doc();
                    }
                });
            }, __("Teams"));

            frm.add_custom_button(__('Cancel Teams Meeting'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.delete_meeting",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        if (r.message) {
                            // If it's an object, show the .message field
                            let msg = (typeof r.message === "string") ? r.message : r.message.message;
                            frappe.msgprint(msg);
                        } else if (r.message && r.message.login_url) {
                            // Redirect to MS login if required
                            window.location.href = r.message.login_url;
                        }
                        frm.reload_doc();
                    }
                });
            }, __("Teams"));
        }
    }
});