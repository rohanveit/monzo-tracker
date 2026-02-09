"""Monzo API client with automatic token refresh."""

from datetime import datetime, timedelta

import requests

from .auth import TokenManager


class MonzoAPI:
    """Wrapper for Monzo API with automatic token refresh on 401."""

    API_URL = 'https://api.monzo.com'

    def __init__(self, token_manager: TokenManager):
        self.token_manager = token_manager

    def _make_request(self, method: str, endpoint: str, **kwargs) -> requests.Response:
        """Make an API request with automatic token refresh on 401."""
        max_retries = 2

        for attempt in range(max_retries):
            access_token = self.token_manager.get_access_token()

            headers = kwargs.pop('headers', {})
            headers['Authorization'] = f'Bearer {access_token}'

            response = requests.request(
                method,
                f"{self.API_URL}{endpoint}",
                headers=headers,
                **kwargs
            )

            if response.status_code == 401:
                print("Token expired or invalid, re-authenticating...")
                self.token_manager.invalidate()
                continue  # Retry with new token

            return response

        raise Exception("Authentication failed after multiple attempts")

    def get_accounts(self) -> list[dict]:
        """Retrieve user's Monzo accounts."""
        response = self._make_request('GET', '/accounts')

        if response.status_code == 200:
            return response.json()['accounts']
        else:
            raise Exception(f"Failed to get accounts: {response.status_code} - {response.text}")

    def get_transactions(self, account_id: str, days: int = 30) -> list[dict]:
        """Retrieve transactions for the last N days, paginating through all results."""
        since = datetime.now() - timedelta(days=days)
        since_str = since.strftime('%Y-%m-%dT%H:%M:%SZ')

        all_transactions = []

        while True:
            params = {
                'account_id': account_id,
                'since': since_str,
                'expand[]': 'merchant',
                'limit': 100,
            }

            response = self._make_request('GET', '/transactions', params=params)

            if response.status_code != 200:
                raise Exception(f"Failed to get transactions: {response.status_code} - {response.text}")

            batch = response.json()['transactions']
            if not batch:
                break

            all_transactions.extend(batch)

            # Use the last transaction's ID as the cursor for the next page
            since_str = batch[-1]['id']

            # If we got fewer than the limit, we've reached the end
            if len(batch) < 100:
                break

        return all_transactions
