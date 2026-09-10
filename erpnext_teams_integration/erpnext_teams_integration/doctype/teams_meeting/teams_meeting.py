# Copyright (c) 2026, Frappe Technologies Pvt. Ltd. and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document
from frappe.contacts.doctype.contact.contact import get_default_contact
from frappe.utils import get_time

class TeamsMeeting(Document):
    def validate(self):
        self.validate_times()

    def before_save(self):
        self.set_participants_email()

    def validate_times(self):
        # Basic sanity check so we don't break the space-time continuum 
        if self.start_time and self.end_time:
            if get_time(self.start_time) >= get_time(self.end_time):
                frappe.throw(_("End Time must be after Start Time. We haven't built a time machine yet!"))

    def add_participant(self, doctype, docname):
        """Add a single participant to meeting participants

        Args:
                doctype (string): Reference Doctype
                docname (string): Reference Docname
        """
        # FIXED: Changed "event_participants" to "meeting_participants"
        self.append(
            "meeting_participants", 
            {
                "reference_doctype": doctype,
                "reference_docname": docname,
            },
        )
    
    def add_participants(self, participants):
        """Add participant entry

        Args:
                participants ([Array]): Array of a dict with doctype and docname
        """
        for participant in participants:
            self.add_participant(participant["doctype"], participant["docname"])

    def set_participants_email(self):
        # Map Doctypes that store emails directly on their own records
        direct_email_map = {
            "Contact": "email_id",
            "Lead": "email_id",
            "User": "email" 
        }

        for participant in self.get("meeting_participants"):
            if participant.email:
                continue

            ref_type = participant.reference_doctype
            ref_name = participant.reference_docname
            email = None

            # Attempt 1: Try grabbing the email directly if it's a mapped Doctype
            if ref_type in direct_email_map:
                email_field = direct_email_map[ref_type]
                email = frappe.get_value(ref_type, ref_name, email_field)

            # Attempt 2: The ultimate fallback - if it's empty or unmapped, hunt down the default contact
            if not email:
                participant_contact = get_default_contact(ref_type, ref_name)
                
                if participant_contact:
                    email = frappe.get_value("Contact", participant_contact, "email_id")
            
            # Apply whatever we found (even if it's still None, at least we tried!)
            participant.email = email