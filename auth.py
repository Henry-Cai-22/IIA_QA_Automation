from msal import ConfidentialClientApplication
from cred import *
import time

"""
This function is used to refresh the user's token.
"""
def refresh_user_token(refresh_token):
    client = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

    # Use the refresh token to get a new access token
    result = client.acquire_token_by_refresh_token(
        refresh_token,
        scopes=SCOPE
    )

    if "access_token" in result:
        new_access_token = result["access_token"]
        expires_in = result["expires_in"]
        return new_access_token, time.time() + expires_in  # New token and expiry time
    return None, None  # If token refresh fails
