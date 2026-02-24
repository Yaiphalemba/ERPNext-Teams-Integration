import frappe
import requests
from werkzeug.wrappers import Response
from datetime import datetime, timedelta
from frappe.utils import now_datetime, cstr
from .helpers import get_settings, get_access_token
import json
import hashlib

@frappe.whitelist(allow_guest=True)
def callback(code=None, state=None, error=None, error_description=None):
    """Handle OAuth callback from Microsoft Teams"""
    
    # Check for OAuth errors first
    if error:
        frappe.log_error(f"OAuth Error: {error} - {error_description}", "Teams OAuth Error")
        frappe.local.response["type"] = "redirect"
        frappe.local.response["location"] = "/app/teams-settings?teams_authentication_status=error"
        return
    
    if not code:
        frappe.throw("Authorization code is missing from callback")
    
    try:
        settings = get_settings()
        
        # Validate required settings
        if not all([settings.client_id, settings.client_secret, settings.tenant_id, settings.redirect_uri]):
            frappe.throw("Teams integration is not properly configured. Please check your settings.")
        
        # Prepare token exchange request
        token_url = f"https://login.microsoftonline.com/{settings.tenant_id}/oauth2/v2.0/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "client_id": settings.client_id,
            "client_secret": settings.client_secret,
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": settings.redirect_uri,
            "scope": "https://graph.microsoft.com/.default"
        }
        
        # Exchange code for tokens
        response = requests.post(token_url, headers=headers, data=data, timeout=30)
        
        if response.status_code != 200:
            error_data = response.json() if response.headers.get('content-type', '').startswith('application/json') else response.text
            frappe.log_error(f"Token exchange failed: {response.status_code} - {error_data}", "Teams Token Exchange Error")
            frappe.throw(f"Failed to authenticate with Microsoft Teams. Please try again.")
        
        token_data = response.json()
        
        # Update settings with new tokens
        settings.access_token = token_data.get("access_token")
        settings.refresh_token = token_data.get("refresh_token")
        
        # Calculate token expiry (subtract 5 minutes for safety buffer)
        expires_in = token_data.get("expires_in", 3600)
        settings.token_expiry = now_datetime() + timedelta(seconds=expires_in - 300)
        
        settings.save(ignore_permissions=True)
        frappe.db.commit()
        
        # Get user info and save Azure ID
        try:
            user_info_response = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers={"Authorization": f"Bearer {settings.access_token}"},
                timeout=30
            )
            
            if user_info_response.status_code == 200:
                user_info = user_info_response.json()
                azure_id = user_info.get("id")
                user_email = user_info.get("mail") or user_info.get("userPrincipalName")
                
                if azure_id:
                    # Update current user's Azure ID
                    if frappe.session.user != "Guest":
                        frappe.db.set_value("User", frappe.session.user, "azure_object_id", azure_id)
                    
                    # Also update based on email if available
                    if user_email and frappe.db.exists("User", {"email": user_email}):
                        frappe.db.set_value("User", {"email": user_email}, "azure_object_id", azure_id)
                    
                    # Update settings with owner info if not set
                    if not settings.azure_owner_email_id and user_email:
                        settings.azure_owner_email_id = user_email
                        settings.owner_azure_object_id = azure_id
                        settings.save(ignore_permissions=True)
                    
                    frappe.db.commit()
                    
        except Exception as e:
            # Log but don't fail the authentication process
            frappe.log_error(f"Failed to fetch user info: {str(e)}", "Teams User Info Error")
        
        # Successful authentication redirect
        redirect_url = "/app/teams-settings?teams_authentication_status=success"
        
        # If state parameter contains redirect info, use it
        if state and state.startswith('from_create_button::'):
            doc_name = state.replace('from_create_button::', '')
            if doc_name:
                if doc_name != "Teams Settings":
                    redirect_url = f"/app/event/{doc_name}?teams_authentication_status=success"
        
        frappe.local.response["type"] = "redirect"
        frappe.local.response["location"] = redirect_url
        
    except Exception as e:
        frappe.log_error(f"Authentication callback error: {str(e)}", "Teams Authentication Error")
        frappe.local.response["type"] = "redirect"
        frappe.local.response["location"] = "/app/teams-settings?teams_authentication_status=error"


@frappe.whitelist()
def get_authentication_status():
    """Check if Teams integration is properly authenticated"""
    try:
        settings = get_settings()
        
        if not settings.access_token:
            return {"authenticated": False, "message": "No access token found"}
        
        # Check if token is expired
        if settings.token_expiry and settings.token_expiry < now_datetime():
            return {"authenticated": False, "message": "Token expired"}
        
        # Test the token by making a simple API call
        headers = {"Authorization": f"Bearer {settings.access_token}"}
        response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers, timeout=10)
        
        if response.status_code == 200:
            return {"authenticated": True, "message": "Authentication successful"}
        else:
            return {"authenticated": False, "message": "Token validation failed"}
            
    except Exception as e:
        frappe.log_error(f"Authentication status check failed: {str(e)}", "Teams Auth Status Error")
        return {"authenticated": False, "message": "Authentication check failed"}


@frappe.whitelist()
def revoke_authentication():
    """Revoke Teams authentication and clear tokens"""
    try:
        settings = get_settings()
        
        # Clear all authentication related fields
        settings.access_token = ""
        settings.refresh_token = ""
        settings.token_expiry = None
        settings.save(ignore_permissions=True)
        frappe.db.commit()
        
        return {"success": True, "message": "Authentication revoked successfully"}
        
    except Exception as e:
        frappe.log_error(f"Failed to revoke authentication: {str(e)}", "Teams Auth Revoke Error")
        frappe.throw("Failed to revoke authentication")

#Webhook and Subscription Management
@frappe.whitelist(allow_guest=True)
def handle_graph_webhook():
    """
    The main listener for Microsoft Graph API Subscriptions.
    Must be a public endpoint (allow_guest=True).
    """
    # 1. The Validation Handshake (Microsoft checking if we are alive)
    validation_token = frappe.request.args.get('validationToken')
    if validation_token:
        # We MUST return plain text, not JSON, or Microsoft rejects the subscription
        frappe.response['type'] = 'text/plain'
        # In some Frappe versions, you have to bypass the formatter entirely:
        frappe.request.environ['werkzeug.request'] = None 
        return Response(validation_token, status=200, mimetype='text/plain')

    # 2. Handling the Actual RSVP Notifications
    payload = frappe.request.get_json()
    
    if payload and "value" in payload:
        for notification in payload.get("value", []):
            resource_url = notification.get("resource")
            
            # Microsoft requires us to respond within 3 seconds, or they assume we failed
            # and will retry aggressively. So, we DO NOT process the RSVP here. 
            # We shove it into a background queue and reply "202 Accepted" immediately.
            frappe.enqueue(
                "erpnext_teams_integration.api.webhooks.process_rsvp_change",
                resource_url=resource_url,
                queue="short"
            )

    # Tell Microsoft we got the message
    frappe.response['type'] = 'text/plain'
    return Response("Accepted", status=202, mimetype='text/plain')

GRAPH_API = "https://graph.microsoft.com/v1.0"

@frappe.whitelist()
def subscribe_to_calendar_events():
    """Tells Microsoft Graph to send webhooks when meetings are updated (RSVPs)."""
    token = get_access_token()
    if not token:
        frappe.throw("Authentication required.")

    # Your public Frappe URL + the API path to your webhook function
    # e.g., "https://my-frappe-site.com/api/method/erpnext_teams_integration.api.webhooks.handle_graph_webhook"
    notification_url = frappe.utils.get_url("/api/method/erpnext_teams_integration.api.webhooks.handle_graph_webhook")

    # Calendar subscriptions expire in maximum 4230 minutes (under 3 days).
    # We set it to expire in 2 days to be safe.
    expiration_time = datetime.utcnow() + timedelta(days=2)

    payload = {
        "changeType": "updated",
        "notificationUrl": notification_url,
        "resource": "/me/events", # Listen to all events on this user's calendar
        "expirationDateTime": expiration_time.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "clientState": "FrappeTeamsSyncV1" # Optional secret to verify it's really MS sending the webhook
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    res = requests.post(f"{GRAPH_API}/subscriptions", headers=headers, json=payload)

    if res.status_code == 201:
        data = res.json()
        # You should ideally save this Subscription ID in your Frappe Settings/DB
        # so you can renew it or delete it later!
        return {"success": True, "subscription_id": data.get("id")}
    else:
        frappe.log_error(f"Subscription failed: {res.text}", "Graph Webhook Error")
        frappe.throw(f"Failed to subscribe: {res.status_code}")
        
def process_rsvp_change(resource_url):
    """
    Background job triggered by Graph API Webhook.
    Fetches the latest event details and updates the Frappe Event Participants table.
    """
    try:
        token = get_access_token()
        if not token: 
            return
            
        # Microsoft sends the resource like "Users('id')/Events('event_id')"
        # We need to build the full API URL
        if not resource_url.startswith("https"):
            url = f"{GRAPH_API}/{resource_url.lstrip('/')}"
        else:
            url = resource_url
            
        headers = {
            "Authorization": f"Bearer {token}", 
            "Content-Type": "application/json"
        }
            
        # 1. Ask Microsoft for the updated Event details
        res = requests.get(url, headers=headers, timeout=30)
        
        if res.status_code != 200:
            frappe.log_error(f"Failed to fetch event RSVP: {res.text}", "RSVP Sync Error")
            return
            
        event_data = res.json()
        outlook_event_id = event_data.get("id")
        attendees = event_data.get("attendees", [])
        
        if not outlook_event_id or not attendees:
            return
            
        # 2. Find the corresponding ERPNext Event Document
        event_name = frappe.db.get_value("Event", {"custom_outlook_event_id": outlook_event_id}, "name")
        if not event_name:
            # We don't track this event, or it was deleted in ERPNext
            return 
            
        # 3. Create a clean lookup dictionary from the Graph API response
        # Graph Statuses: 'none', 'tentative', 'accepted', 'declined', 'notResponded'
        status_map = {
            "accepted": "Yes",
            "declined": "No",
            "tentative": "Maybe" # If you don't have 'Maybe' in your Select field, map this to blank or remove it
        }
        
        attendee_responses = {}
        for a in attendees:
            email = a.get("emailAddress", {}).get("address", "").lower()
            status = a.get("status", {}).get("response", "").lower()
            
            if email and status in status_map:
                attendee_responses[email] = status_map[status]
                
        # 4. Map the responses to the ERPNext Child Table
        doc = frappe.get_doc("Event", event_name)
        doc_updated = False
        
        for row in doc.event_participants:
            row_email = (row.email or "").lower()
            
            if row_email in attendee_responses:
                new_status = attendee_responses[row_email]
                
                # Only update if the status actually changed to avoid unnecessary saves
                if row.attending != new_status:
                    row.attending = new_status
                    doc_updated = True
                    
        # 5. Save the document silently
        if doc_updated:
            # Prevent circular loops if you have "on_update" hooks that send updates back to Teams
            doc.flags.ignore_validate = True 
            doc.flags.ignore_permissions = True
            doc.save()
            frappe.db.commit()
            
    except Exception as e:
        frappe.log_error(f"Error processing RSVP: {str(e)}", "RSVP Processing Error")
        
        
def renew_graph_subscriptions():
    """
    Cron job to keep the Microsoft Graph Webhook alive.
    Runs daily.
    """
    try:
        token = get_access_token()
        if not token: 
            return
            
        settings = frappe.get_single("Teams Settings")
        sub_id = settings.get("custom_webhook_subscription_id")
        
        headers = {
            "Authorization": f"Bearer {token}", 
            "Content-Type": "application/json"
        }
        
        # If we have an existing subscription, try to renew it
        if sub_id:
            # Extend for another 48 hours
            expiration_time = datetime.utcnow() + timedelta(days=2)
            payload = {
                "expirationDateTime": expiration_time.strftime("%Y-%m-%dT%H:%M:%SZ")
            }
            
            res = requests.patch(f"{GRAPH_API}/subscriptions/{sub_id}", headers=headers, json=payload)
            
            if res.status_code == 200:
                # Successfully renewed, we're done here!
                return
            else:
                # 404 means it already expired and Microsoft deleted it. We must recreate.
                frappe.log_error(f"Failed to renew webhook (likely expired). Recreating... {res.text}", "Webhook Renewal")
        
        # If we reach here, we either didn't have a sub_id, or the renewal failed.
        # Time to create a brand new one. (Assuming you created the subscribe function from the previous step)
        result = subscribe_to_calendar_events()
        
        if result and result.get("success"):
            # Save the new ID so we can renew it tomorrow
            settings.db_set("custom_webhook_subscription_id", result.get("subscription_id"))
            frappe.db.commit()
            
    except Exception as e:
        # 1. Grab the full error string for the body
        error_details = str(e)
        
        # 2. Slice the title to strictly 135 characters just to be safe
        safe_title = f"Webhook Renewal Error: {error_details}"[:135]
        
        # 3. Log it cleanly
        frappe.log_error(message=frappe.get_traceback(), title=safe_title)