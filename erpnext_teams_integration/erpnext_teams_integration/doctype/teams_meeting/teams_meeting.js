// Copyright (c) 2026, Frappe Technologies Pvt. Ltd. and contributors
// For license information, please see license.txt

frappe.provide("frappe.desk");

frappe.ui.form.on("Teams Meeting", {
    onload: function (frm) {
        // Restrict participant doctypes to the ones that actually make sense
        frm.set_query("reference_doctype", "meeting_participants", function () {
            return {
                filters: {
                    name: ["in", [
                        "User", 
                        "Lead", 
                        "Customer", 
                        "Contact", 
                        "Supplier", 
                        "Sales Partner"
                    ]]
                }
            };
        });
    },

    refresh: function (frm) {
        // Render a sleek dashboard headline if the meeting URL exists and it's not cancelled
        if (frm.doc.teams_meeting_url && frm.doc.status !== "Cancelled") {
            frm.dashboard.set_headline(
                __(`Ready to jump in? <a target='_blank' href='${frm.doc.teams_meeting_url}' class='btn btn-primary btn-xs' style='margin-left: 10px;'>Join Teams Meeting</a>`)
            );
        }

        // Add custom buttons for quick access to participant records
        if (frm.doc.meeting_participants) {
            frm.doc.meeting_participants.forEach((value) => {
                if (value.reference_docname) {
                    frm.add_custom_button(
                        __(value.reference_docname),
                        function () {
                            frappe.set_route("Form", value.reference_doctype, value.reference_docname);
                        },
                        __("Participants")
                    );
                }
            });
        }

        if (frm.doc.custom_teams_meeting_url) {
            frm.fields_dict.custom_join_teams_meeting.$wrapper
                .find("button")
                .off("click") // Remove existing handler if any
                .on("click", function () {
                    window.open(frm.doc.custom_teams_meeting_url, "_blank");
                });
        }

        // --- TEAMS INTEGRATION LOGIC ---

        // 1. Check for successful authentication token in the URL and clear it[cite: 2]
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

        // 2. Add API action buttons if the document is saved[cite: 2]
        if (!frm.doc.__islocal) {
            
            // ==========================================
            // 1. CREATE MEETING
            // ==========================================
            frm.add_custom_button(__('Create Teams Meeting'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.create_meeting",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        // Auth check first
                        if (r.message && r.message.login_url) {
                            window.location.href = r.message.login_url;
                            return;
                        }
                        
                        let msg = (typeof r.message === "string") 
                            ? r.message 
                            : (r.message?.message || __("Meeting created successfully."));
                        
                        // For Create, we usually just need to reload to fetch the new meeting links/IDs
                        frm.reload_doc()
                            .then(() => {
                                frm.set_value("status", "Open"); // Or "Open" depending on your workflow
                                return frm.save();
                            })
                            .then(() => {
                                frappe.show_alert({ message: msg, indicator: "green" });
                            })
                            .catch((err) => console.error("Error updating document:", err));
                    }
                });
            }, __("Teams"));


            // ==========================================
            // 2. RESCHEDULE MEETING
            // ==========================================
            frm.add_custom_button(__('Reschedule Teams Meeting'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.reschedule_meeting",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        // Auth check first
                        if (r.message && r.message.login_url) {
                            window.location.href = r.message.login_url;
                            return;
                        }
                        
                        let msg = (typeof r.message === "string") 
                            ? r.message 
                            : (r.message?.message || __("Meeting rescheduled successfully."));
                        
                        // Reload -> Set Value -> Save
                        frm.reload_doc()
                            .then(() => {
                                frm.set_value("status", "Rescheduled"); // Or "Open" depending on your workflow
                                return frm.save();
                            })
                            .then(() => {
                                frappe.show_alert({ message: msg, indicator: "green" });
                            })
                            .catch((err) => console.error("Error updating document:", err));
                    }
                });
            }, __("Teams"));


            // ==========================================
            // 3. CANCEL MEETING
            // ==========================================
            frm.add_custom_button(__('Cancel Teams Meeting'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.delete_meeting",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        // Auth check first
                        if (r.message && r.message.login_url) {
                            window.location.href = r.message.login_url;
                            return;
                        }
                        
                        let msg = (typeof r.message === "string") 
                            ? r.message 
                            : (r.message?.message || __("Meeting cancelled successfully."));
                        
                        // Reload -> Set Value -> Save
                        frm.reload_doc()
                            .then(() => {
                                frm.set_value("status", "Cancelled"); // Or "Open"
                                return frm.save();
                            })
                            .then(() => {
                                frappe.show_alert({ message: msg, indicator: "green" });
                            })
                            .catch((err) => console.error("Error updating document:", err));
                    }
                });
            }, __("Teams"));

            // ==========================================
            // 3. GET MEETING RECORDINGS
            // ==========================================
            frm.add_custom_button(__('Get Teams Meeting Recording'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.fetch_meeting_recording",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        if (r.message) {
                            console.log(r);
                            frappe.msgprint("Meeting recording fetched successfully. Please check the 'Meeting Recordings' table for the recording URLs.");
                        } else if (r.message && r.message.login_url) {
                            window.location.href = r.message.login_url;
                        }
                        frm.reload_doc();
                    },
                    error: function(err) {
                        console.error("Error fetching meeting recording:", err);
                    }
                });
            }, __("Teams"));
        }
    },

    start_time: function (frm) {
        // End Time is auto-populated to 30 minutes after the Start Time if End Time is not set
        if (frm.doc.start_time && !frm.doc.end_time) {
            // Use moment.js to handle the time math safely
            let new_time = moment(frm.doc.start_time, "HH:mm:ss")
                .add(30, 'minutes')
                .format("HH:mm:ss");
            
            frm.set_value("end_time", new_time);
        }
    }
});

frappe.desk.meeting_participantsParticipants = class meetingParticipants {
    constructor(frm, doctype) {
        this.frm = frm;
        this.doctype = doctype;
        this.make();
    }

    make() {
        let me = this;
        let table = me.frm.get_field("meeting_participants").grid; 
        
        new frappe.ui.form.LinkSelector({
            doctype: me.doctype,
            dynamic_link_field: "reference_doctype",
            dynamic_link_reference: me.doctype,
            fieldname: "reference_docname",
            target: table,
            txt: "",
        });
    }
};

frappe.ui.form.on("Meeting Recordings", {
    playdownload(frm, cdt, cdn) {
        let row = locals[cdt][cdn];
        let api_method = "erpnext_teams_integration.api.meetings.stream_meeting_recording";
        let target_url = row.recording_url;
        
        setTimeout(() => {
            let download_url = `/api/method/${api_method}?docname=${encodeURIComponent(frm.doc.name)}&doctype=${encodeURIComponent(frm.doc.doctype)}&target_url=${encodeURIComponent(target_url.trim())}&index=${row.idx}`;
            window.open(download_url, '_blank');
        }, row.idx * 1500); // 1.5 second delay between each file
    }
})