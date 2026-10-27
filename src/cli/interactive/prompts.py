"""
CLI Prompt Utilities
====================
Provides clean, user-friendly interactive prompts with smart defaults.
"""

from typing import List, Optional


def ask_string(prompt: str, default: str = "") -> str:
    """
    Prompts the user for a text string, displaying the default value in brackets.
    """
    default_hint = f" [Default: {default}]" if default else ""
    user_input = input(f"{prompt}{default_hint}: ").strip()
    return user_input if user_input else default


def ask_choice(prompt: str, valid_choices: List[str], default: str) -> str:
    """
    Prompts the user to select from a list of predefined choices.
    """
    choices_str = "/".join(valid_choices)
    while True:
        user_input = input(f"{prompt} ({choices_str}) [Default: {default}]: ").strip().upper()
        if not user_input:
            return default
        if user_input in [c.upper() for c in valid_choices]:
            return user_input
        print(f"[!] Invalid option '{user_input}'. Please choose from: {choices_str}")


def ask_int(prompt: str, default: int) -> int:
    """
    Prompts the user for an integer value with validation.
    """
    while True:
        user_input = input(f"{prompt} [Default: {default}]: ").strip()
        if not user_input:
            return default
        try:
            return int(user_input)
        except ValueError:
            print("[!] Please enter a valid integer.")
