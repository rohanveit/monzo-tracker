"""Data models and formatting for Monzo transactions."""

from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Field, computed_field


class Merchant(BaseModel):
    """Merchant information from a transaction."""

    id: Optional[str] = None
    name: Optional[str] = None
    category: Optional[str] = None
    logo: Optional[str] = None


class Transaction(BaseModel):
    """Raw transaction from Monzo API."""

    id: str
    amount: int  # Amount in pence
    currency: str
    created: datetime
    description: str = ""
    category: str = "unknown"
    notes: str = ""
    merchant: Optional[Merchant] = None

    @computed_field
    @property
    def amount_pounds(self) -> float:
        """Amount converted to pounds."""
        return self.amount / 100

    @computed_field
    @property
    def display_description(self) -> str:
        """Human-readable description (merchant name or fallback)."""
        if self.merchant and self.merchant.name:
            return self.merchant.name
        return self.description or "Unknown"


class FormattedTransaction(BaseModel):
    """Formatted transaction for display."""

    id: str
    date: str
    description: str
    amount: str
    amount_raw: float
    currency: str
    category: str
    notes: str

    @classmethod
    def from_transaction(cls, tx: Transaction) -> "FormattedTransaction":
        """Create a formatted transaction from a raw transaction."""
        currency = tx.currency.upper()
        return cls(
            id=tx.id,
            date=tx.created.strftime("%Y-%m-%d %H:%M:%S"),
            description=tx.display_description,
            amount=f"{currency} {tx.amount_pounds:.2f}",
            amount_raw=tx.amount_pounds,
            currency=currency,
            category=tx.category,
            notes=tx.notes,
        )


def format_transaction(tx: dict) -> FormattedTransaction:
    """Format a raw transaction for display.

    Args:
        tx: Raw transaction dict from Monzo API

    Returns:
        FormattedTransaction model instance
    """
    # Parse the raw dict into a Transaction model
    raw_tx = tx.copy()
    # Handle the Z suffix in ISO format
    if isinstance(raw_tx.get("created"), str):
        raw_tx["created"] = raw_tx["created"].replace("Z", "+00:00")

    transaction = Transaction.model_validate(raw_tx)
    return FormattedTransaction.from_transaction(transaction)
