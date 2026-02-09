"""OAuth authentication and token management for Monzo API."""

import json
import os
import time
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import requests

# Token storage
TOKEN_FILE = Path.home() / '.monzo_tokens.json'
TOKEN_EXPIRY_BUFFER = 300  # 5 minutes buffer before expiration


class OAuthCallbackHandler(BaseHTTPRequestHandler):
    """Handler for OAuth callback."""

    def do_GET(self):
        """Handle the OAuth callback."""
        query = urlparse(self.path).query
        params = parse_qs(query)

        if 'code' in params:
            self.server.auth_code = params['code'][0]
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(
                b'<html><body><h1>Authentication successful!</h1>'
                b'<p>You can close this window.</p></body></html>'
            )
        else:
            self.send_response(400)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            self.wfile.write(b'<html><body><h1>Authentication failed!</h1></body></html>')

    def log_message(self, format, *args):
        """Suppress log messages."""
        pass


def start_oauth_flow(client_id: str, redirect_uri: str, auth_url: str) -> str | None:
    """Initiate OAuth flow and return authorization code."""
    auth_params = {
        'client_id': client_id,
        'redirect_uri': redirect_uri,
        'response_type': 'code',
        'state': 'random_state_string'  # In production, use a secure random string
    }

    full_auth_url = f"{auth_url}/?{'&'.join([f'{k}={v}' for k, v in auth_params.items()])}"

    print("Opening browser for authentication...")
    print(f"If browser doesn't open, visit: {full_auth_url}")

    webbrowser.open(full_auth_url)

    # Start local server to receive callback
    server = HTTPServer(('localhost', 8080), OAuthCallbackHandler)
    server.auth_code = None

    print("Waiting for authentication callback...")
    server.handle_request()

    return server.auth_code


def exchange_code_for_token(
    auth_code: str,
    client_id: str,
    client_secret: str,
    redirect_uri: str,
    api_url: str
) -> dict:
    """Exchange authorization code for access token."""
    token_url = f"{api_url}/oauth2/token"

    data = {
        'grant_type': 'authorization_code',
        'client_id': client_id,
        'client_secret': client_secret,
        'redirect_uri': redirect_uri,
        'code': auth_code
    }

    response = requests.post(token_url, data=data)

    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f"Token exchange failed: {response.status_code} - {response.text}")


class TokenManager:
    """Manages OAuth tokens with persistence and automatic refresh."""

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        redirect_uri: str,
        auth_url: str = 'https://auth.monzo.com',
        api_url: str = 'https://api.monzo.com'
    ):
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri
        self.auth_url = auth_url
        self.api_url = api_url
        self.token_data = None
        self._load_tokens()

    def _load_tokens(self):
        """Load tokens from disk if they exist."""
        if TOKEN_FILE.exists():
            try:
                with open(TOKEN_FILE, 'r') as f:
                    self.token_data = json.load(f)
            except (json.JSONDecodeError, IOError):
                self.token_data = None

    def _save_tokens(self, token_response: dict):
        """Save token response to disk with calculated expiry time."""
        self.token_data = {
            'access_token': token_response['access_token'],
            'refresh_token': token_response.get('refresh_token'),
            'token_type': token_response.get('token_type', 'Bearer'),
            'expires_at': time.time() + token_response.get('expires_in', 21600),
            'user_id': token_response.get('user_id')
        }

        with open(TOKEN_FILE, 'w') as f:
            json.dump(self.token_data, f, indent=2)

        # Set file permissions to owner-only (Unix/Linux/Mac)
        try:
            os.chmod(TOKEN_FILE, 0o600)
        except OSError:
            pass  # Windows doesn't support this

    def is_token_valid(self) -> bool:
        """Check if current token is valid and not expired."""
        if not self.token_data or not self.token_data.get('access_token'):
            return False

        expires_at = self.token_data.get('expires_at', 0)
        return time.time() < (expires_at - TOKEN_EXPIRY_BUFFER)

    def get_access_token(self) -> str:
        """Get a valid access token, refreshing or re-authenticating as needed."""
        # Try existing token first
        if self.is_token_valid():
            return self.token_data['access_token']

        # Try refresh token if available
        if self.token_data and self.token_data.get('refresh_token'):
            try:
                return self._refresh_token()
            except Exception as e:
                print(f"Token refresh failed: {e}")
                # Fall through to full re-auth

        # Full re-authentication required
        print("Authentication required. Starting OAuth flow...")
        return self._full_authentication()

    def _refresh_token(self) -> str:
        """Attempt to refresh the access token using refresh_token."""
        print("Refreshing access token...")

        data = {
            'grant_type': 'refresh_token',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'refresh_token': self.token_data['refresh_token']
        }

        response = requests.post(f"{self.api_url}/oauth2/token", data=data)

        if response.status_code == 200:
            token_response = response.json()
            self._save_tokens(token_response)
            print("Token refreshed successfully")
            return self.token_data['access_token']
        else:
            raise Exception(f"Refresh failed: {response.status_code} - {response.text}")

    def _full_authentication(self) -> str:
        """Perform full OAuth flow."""
        auth_code = start_oauth_flow(self.client_id, self.redirect_uri, self.auth_url)
        if not auth_code:
            raise Exception("Failed to get authorization code")

        token_response = exchange_code_for_token(
            auth_code,
            self.client_id,
            self.client_secret,
            self.redirect_uri,
            self.api_url
        )
        self._save_tokens(token_response)

        print("Authentication successful!")
        print("IMPORTANT: Open your Monzo app and approve this login.")
        input("Press Enter once you've approved in the Monzo app...")
        return self.token_data['access_token']

    def invalidate(self):
        """Invalidate current tokens (call on 401 error)."""
        if self.token_data:
            self.token_data['expires_at'] = 0  # Force expiration
