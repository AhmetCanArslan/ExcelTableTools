import re
import json
import os
from pathlib import Path
from urllib.request import urlopen
from datetime import datetime, timedelta
import logging

class DomainValidator:
    PSL_URL = "https://publicsuffix.org/list/public_suffix_list.dat"
    CACHE_DURATION_DAYS = 30
    
    def __init__(self, config_dir=None):
        self.config_dir = config_dir or Path(__file__).parent.parent / 'config'
        self.config_dir.mkdir(exist_ok=True)
        self.psl_cache_file = self.config_dir / 'public_suffix_list.dat'
        self.custom_domains_file = self.config_dir / 'custom_domains.json'
        self.domain_patterns = {
            r'\.edu$': 'Educational Institution',
            r'\.edu\.[a-z]{2}$': 'International Educational Institution',
            r'\.ac\.[a-z]{2}$': 'Academic Institution',
            r'\.gov$': 'Government Institution',
            r'\.gov\.[a-z]{2}$': 'International Government Institution',
            r'\.mil$': 'Military Institution',
            r'\.org$': 'Organization',
            r'\.com$': 'Commercial',
            r'\.net$': 'Network',
        }
        
        # Initialize custom domains
        self._init_custom_domains()
        # Load or update PSL
        self._load_or_update_psl()
    
    def _init_custom_domains(self):
        """Initialize or load custom domains from JSON file."""
        if not self.custom_domains_file.exists():
            self.custom_domains = {
                "allowed": set(),
                "blocked": set()
            }
            self._save_custom_domains()
        else:
            self._load_custom_domains()
    
    def _load_custom_domains(self):
        """Load custom domains from JSON file."""
        try:
            with open(self.custom_domains_file, 'r') as f:
                data = json.load(f)
                self.custom_domains = {
                    "allowed": set(data.get("allowed", [])),
                    "blocked": set(data.get("blocked", []))
                }
        except Exception as e:
            logging.error(f"Error loading custom domains: {e}")
            self.custom_domains = {"allowed": set(), "blocked": set()}
    
    def _save_custom_domains(self):
        """Save custom domains to JSON file."""
        try:
            with open(self.custom_domains_file, 'w') as f:
                json.dump({
                    "allowed": list(self.custom_domains["allowed"]),
                    "blocked": list(self.custom_domains["blocked"])
                }, f, indent=2)
        except Exception as e:
            logging.error(f"Error saving custom domains: {e}")
    
    def _load_or_update_psl(self):
        """Load or update the Public Suffix List."""
        need_update = True
        if self.psl_cache_file.exists():
            mtime = datetime.fromtimestamp(self.psl_cache_file.stat().st_mtime)
            if datetime.now() - mtime < timedelta(days=self.CACHE_DURATION_DAYS):
                need_update = False
        
        if need_update:
            try:
                with urlopen(self.PSL_URL) as response:
                    content = response.read().decode('utf-8')
                    # Filter out comments and empty lines
                    valid_lines = [line.strip() for line in content.splitlines()
                                if line.strip() and not line.startswith('//')]
                    with open(self.psl_cache_file, 'w') as f:
                        f.write('\n'.join(valid_lines))
            except Exception as e:
                logging.error(f"Error updating PSL: {e}")
                if not self.psl_cache_file.exists():
                    # Create minimal PSL if download fails and no cache exists
                    with open(self.psl_cache_file, 'w') as f:
                        f.write('\n'.join([
                            "com", "org", "net", "edu", "gov", "mil",
                            "*.com", "*.org", "*.net", "*.edu", "*.gov", "*.mil"
                        ]))
        
        # Load PSL into memory
        with open(self.psl_cache_file, 'r') as f:
            self.public_suffixes = set(line.strip() for line in f)
    
    def add_custom_domain(self, domain, is_allowed=True):
        """Add a domain to custom allowed or blocked list."""
        domain = domain.lower()
        if is_allowed:
            self.custom_domains["blocked"].discard(domain)
            self.custom_domains["allowed"].add(domain)
        else:
            self.custom_domains["allowed"].discard(domain)
            self.custom_domains["blocked"].add(domain)
        self._save_custom_domains()
    
    def remove_custom_domain(self, domain):
        """Remove a domain from custom lists."""
        domain = domain.lower()
        self.custom_domains["allowed"].discard(domain)
        self.custom_domains["blocked"].discard(domain)
        self._save_custom_domains()
    
    def is_valid_domain(self, domain):
        """
        Validate a domain using multiple checks.
        Returns (is_valid, reason) tuple.
        """
        domain = domain.lower()
        
        # Check custom lists first
        if domain in self.custom_domains["blocked"]:
            return False, "Blocked Domain"
        if domain in self.custom_domains["allowed"]:
            return True, "Allowed Domain"
        
        # Check domain format
        if not re.match(r'^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)*$', domain):
            return False, "Invalid Format"
        
        # Check domain parts
        parts = domain.split('.')
        if len(parts) < 2:
            return False, "Invalid Format"
        
        # Check against PSL
        suffix = '.'.join(parts[-2:])  # Consider last two parts as potential TLD
        if suffix in self.public_suffixes or f"*.{parts[-1]}" in self.public_suffixes:
            # Check for institutional patterns
            for pattern, domain_type in self.domain_patterns.items():
                if re.search(pattern, domain):
                    return True, f"Valid {domain_type}"
            return True, "Valid Domain"
        
        return False, "Unknown TLD"

def validate_email_address(email, validator=None):
    """Validate an email address and its domain."""
    if validator is None:
        validator = DomainValidator()
    
    # Basic email format validation
    email = str(email).strip()
    email_pattern = r"^(?!.*\.\.)(?!.*\.$)[^\W][\w.%+-]{0,63}@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    if not re.match(email_pattern, email):
        return False, "Invalid Email Format"
    
    # Extract and validate domain
    domain = email.split('@')[-1].lower()
    return validator.is_valid_domain(domain)

def print_validation_result(email, is_valid, reason):
    """Print validation result in a formatted way."""
    print("\nEmail Validation Result:")
    print("-" * 50)
    print(f"Email:  {email}")
    print(f"Status: {'✓ Valid' if is_valid else '✗ Invalid'}")
    print(f"Reason: {reason}")
    print("-" * 50)

def main():
    """Command-line interface for email domain validation."""
    print("\nEmail Domain Validator")
    print("=" * 50)
    print("This tool validates email addresses and their domains.")
    print("Type 'quit' or 'exit' to end the program.")
    print("=" * 50)
    
    # Initialize validator once to reuse
    validator = DomainValidator()
    
    while True:
        try:
            email = input("\nEnter an email address to validate: ").strip()
            
            if email.lower() in ('quit', 'exit'):
                print("\nGoodbye!")
                break
            
            if not email:
                print("Please enter an email address.")
                continue
            
            is_valid, reason = validate_email_address(email, validator)
            print_validation_result(email, is_valid, reason)
            
        except KeyboardInterrupt:
            print("\n\nTerminating program...")
            break
        except Exception as e:
            print(f"\nError: {str(e)}")

if __name__ == '__main__':
    main() 