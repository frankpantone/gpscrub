"""
Hertz PDF Scrubber and Email Generator
Extracts vehicle info, location, and Hertz rep emails from gate release PDFs
and creates Outlook email drafts for logistics confirmation.
"""

import pypdf
import re
import win32com.client
import sys
from datetime import datetime, timedelta
from pathlib import Path


class HertzPDFProcessor:
    """Process Hertz gate release PDFs and extract relevant information."""
    
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.text = ""
        self.data = {}
        
    def extract_text(self):
        """Extract text from PDF."""
        with open(self.pdf_path, 'rb') as file:
            reader = pypdf.PdfReader(file)
            self.text = ""
            for page in reader.pages:
                self.text += page.extract_text()
        return self.text
    
    def parse_data(self):
        """Parse and extract structured data from PDF text."""
        if not self.text:
            self.extract_text()
        
        # Extract email addresses with their positions
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        all_emails = re.findall(email_pattern, self.text)
        
        # Filter Hertz representative emails (exclude purchaser emails)
        # Strategy: Emails in Instruction field or with "Kori Novak" are key indicators
        hertz_emails = []
        excluded_emails = []
        text_lower = self.text.lower()
        
        # HARDCODED EXCLUSION: Always exclude konovak@hertz.com (purchaser email)
        ALWAYS_EXCLUDE = ['konovak@hertz.com']
        
        # Find key indicators
        kori_novak_pos = text_lower.find('kori novak')
        
        # Look for contact name patterns (e.g., "Name1/Name2" or "Name1 / Name2")
        contact_pattern = r'([A-Z][a-z]+\s+[A-Z][a-z]+)\s*/\s*([A-Z][a-z]+\s+[A-Z][a-z]+)'
        contact_matches = re.findall(contact_pattern, self.text)
        
        # Extract names from contact patterns (these are likely Hertz reps)
        hertz_rep_names = []
        for name1, name2 in contact_matches:
            hertz_rep_names.extend([name1.lower(), name2.lower()])
        
        # Get all Hertz emails with their positions
        email_positions = []
        for email in all_emails:
            if '@hertz.com' in email.lower():
                pos = text_lower.find(email.lower())
                email_positions.append((pos, email))
        
        email_positions.sort()  # Sort by position
        
        # First pass: identify emails associated with Hertz rep names
        hertz_rep_emails_by_name = []
        for pos, email in email_positions:
            # Check if any Hertz rep name appears near this email (within 500 chars)
            for rep_name in hertz_rep_names:
                if abs(pos - text_lower.find(rep_name)) < 500:
                    hertz_rep_emails_by_name.append(email)
                    break
        
        # Find the email closest to "Kori Novak" that's NOT a Hertz rep by name
        purchaser_email = None
        if kori_novak_pos != -1:
            min_distance = float('inf')
            for pos, email in email_positions:
                if email not in hertz_rep_emails_by_name:  # Skip Hertz rep emails
                    distance = abs(pos - kori_novak_pos)
                    if distance < min_distance and distance < 400:  # Within reasonable range
                        min_distance = distance
                        purchaser_email = email
        
        # Second pass: categorize all emails
        for pos, email in email_positions:
            # FIRST: Check hardcoded exclusion list
            if email.lower() in [e.lower() for e in ALWAYS_EXCLUDE]:
                excluded_emails.append(email)
            elif email == purchaser_email:
                # This is the purchaser email (closest to Kori Novak)
                excluded_emails.append(email)
            else:
                # All other emails are Hertz reps
                hertz_emails.append(email)
        
        self.data['hertz_rep_emails'] = list(set(hertz_emails))
        self.data['all_emails'] = list(set(all_emails))
        self.data['excluded_emails'] = list(set(excluded_emails))
        
        # Extract VIN (Vehicle Identification Number) - 17 character alphanumeric
        vin_pattern = r'\b[A-HJ-NPR-Z0-9]{17}\b'
        vin_match = re.search(vin_pattern, self.text)
        self.data['vin'] = vin_match.group(0) if vin_match else "VIN_NOT_FOUND"
        
        # Extract vehicle information
        # Year (4 digits, typically 19xx or 20xx, but avoid dates like 2025-10-28)
        # Look for standalone years not part of a date
        year_pattern = r'(?<!\d)(19\d{2}|20[0-2]\d)(?!\d|-)' 
        year_matches = re.findall(year_pattern, self.text)
        # Filter to reasonable model year range (current year down to 1990)
        current_year = datetime.now().year
        valid_years = [y for y in year_matches if 1990 <= int(y) <= current_year]
        # Model year typically appears earlier in the doc than pickup dates
        self.data['year'] = valid_years[0] if valid_years else "YEAR"
        
        # Common car makes (expand this list as needed)
        makes = ['CHEVROLET', 'FORD', 'TOYOTA', 'HONDA', 'KIA', 'HYUNDAI', 'NISSAN', 
                 'MAZDA', 'SUBARU', 'VOLKSWAGEN', 'BMW', 'MERCEDES', 'AUDI', 'JEEP',
                 'RAM', 'GMC', 'CADILLAC', 'LEXUS', 'ACURA', 'INFINITI', 'BUICK',
                 'CHRYSLER', 'DODGE', 'LINCOLN', 'TESLA', 'VOLVO', 'GENESIS', 'ALFA',
                 'POLESTAR', 'RIVIAN', 'LUCID', 'MITSUBISHI']
        
        # Search for makes - prioritize those NOT in company names or email addresses
        found_makes = []
        text_upper = self.text.upper()
        text_lines = text_upper.split('\n')
        
        for make in makes:
            # Use word boundary search to avoid matching makes inside other words
            pattern = r'\b' + make + r'\b'
            match = re.search(pattern, text_upper)
            
            if match:
                pos = match.start()
                
                # Check if the make is actually INSIDE an email address
                # Look for pattern: <chars>@<make or chars-with-make>.<domain>
                in_email = False
                context_start = max(0, pos - 50)
                context_end = min(len(text_upper), pos + len(make) + 20)
                context = text_upper[context_start:context_end]
                
                # Check if make appears between @ and a dot (like hartford@...com)
                # or appears right before @ (like ford@hertz.com)
                email_context = context[max(0, pos - context_start - 20):pos - context_start + len(make) + 20]
                if re.search(r'[A-Z0-9]+@[A-Z0-9]*' + make, email_context) or \
                   re.search(make + r'[A-Z0-9]*@', email_context):
                    in_email = True
                
                # Check if it's in a dealership/company name line
                in_company_name = False
                for line in text_lines:
                    if make in line and any(word in line for word in ['DEALERSHIP', 'INC', 'COMPANY', 'CORP', 'LLC']):
                        in_company_name = True
                        break
                
                # Skip makes in email addresses, deprioritize makes in company names
                if in_email:
                    continue  # Skip entirely
                
                priority = pos + (10000 if in_company_name else 0)
                found_makes.append((priority, make))
        
        # Sort by priority (position + penalty) and take the best match
        if found_makes:
            found_makes.sort(key=lambda x: x[0])
            self.data['make'] = found_makes[0][1].title()
        else:
            self.data['make'] = "MAKE"
        
        # Extract model (look for common patterns near the make)
        # Match both all-letter models (SPORTAGE) and alphanumeric models (K4, K5, Q5, X3)
        model_pattern = r'\b([A-Z][A-Z0-9]+)\b'
        models = re.findall(model_pattern, self.text)
        # Filter out common non-model words including city names
        exclude_words = {'HERTZ', 'INDIANAPOLIS', 'QUINCY', 'GRAY', 'BLACK', 'WHITE', 
                        'RED', 'BLUE', 'SILVER', 'FWD', 'AWD', 'RWD', 'AUTO', 'MANUAL',
                        'TRANSPORTER', 'DRIVER', 'PURCHASER', 'INFORMATION', 'VEHICLE',
                        'SIGNATURE', 'AUTHORIZED', 'DEALERSHIP', 'COMPANY', 'TRANSPORT',
                        'PRESENTED', 'RELEASE', 'PICKED', 'LICENSE', 'NUMBER', 'DATE',
                        'NOTED', 'DAMAGE', 'AREA', 'UNIT', 'SERIAL', 'ODOMETER', 'MODEL',
                        'YEAR', 'COLOR', 'MAKE', 'BODY', 'TYPE', 'LOCATION', 'GATE',
                        'CONTACT', 'PHONE', 'ADDRESS', 'EMAIL', 'HOURS', 'OPERATION',
                        'ORDER', 'INSTRUCTION', 'REPRESENTATIVE', 'MANAGER', 'OFFICE',
                        # Add transmission and body types
                        'CVT', 'IVT', 'NATL', 'SDN', 'SEDAN', 'COUPE', 'SUV', 'WAGON', 'HATCHBACK',
                        'CONVERTIBLE', 'PICKUP', 'TRUCK', 'VAN', 'MINIVAN',
                        # Add trim levels and engine designations
                        'B4', 'B5', 'B6', 'T4', 'T5', 'T6', 'T8', 'S60', 'S90', 'V60', 'V90',
                        'AWD', 'RWD', 'FWD', 'HYB', 'PHEV', 'HYBRID',
                        # Add truck sizes and designations
                        'MED', 'MEDIUM', 'HEAVY', 'LIGHT', 'RF', 'GVWR',
                        # Add location codes and other abbreviations
                        'AP', 'RHP', 'TITLE', 'RETAIL', 'EXT', 'INT', 'STD', 'AM', 'PM',
                        # Add common name parts that might get picked up
                        'IBRAHIM', 'DESOTELL', 'NOVAK', 'SWIFT', 'GOUDEAU',
                        # Add airport codes and location abbreviations
                        'JFK', 'LAX', 'ORD', 'DFW', 'ATL', 'DEN', 'SFO', 'SEA', 'LAS',
                        'MCO', 'PHX', 'IAH', 'MIA', 'BOS', 'EWR', 'MSP', 'DTW', 'PHL',
                        # Add make names that might appear in dealership names
                        'GMC', 'CDJR', 'CHRYSLER',
                        # Add common city names and states
                        'VEGAS', 'ANGELES', 'FRANCISCO', 'DIEGO', 'ANTONIO', 'JOSE',
                        'AUSTIN', 'DALLAS', 'HOUSTON', 'CHICAGO', 'PHOENIX', 'SEATTLE',
                        'BOSTON', 'ATLANTA', 'MIAMI', 'DENVER', 'PORTLAND', 'DETROIT',
                        'MEMPHIS', 'NASHVILLE', 'BALTIMORE', 'MILWAUKEE', 'CHARLOTTE',
                        'COLUMBUS', 'CLEVELAND', 'PITTSBURGH', 'CINCINNATI', 'KANSAS',
                        'LOUIS', 'TAMPA', 'ORLANDO', 'SACRAMENTO', 'RALEIGH', 'NEWARK',
                        'JERSEY', 'CLARA', 'YORK', 'ISLAND', 'CITY', 'PRESENT',
                        # Add other common words (removed SANTA to allow Santa Fe)
                        'FLEET', 'GROUP', 'TEAM', 'MANAGER', 'DISTRIBUTION', 'POOL',
                        'LOGISTICS', 'RENTAL', 'COUNTER', 'AIRPORT', 'DOWNTOWN',
                        # Add US state codes (2-letter)
                        'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA',
                        'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
                        'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
                        'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
                        'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY',
                        'DC',  # Washington DC
                        # HARDCODED EXCLUSION: Never allow these as model names
                        'INC', 'ST', 'LLC', 'CORP', 'LTD', 'CO', 'HWY', 'HIGHWAY'}  # Business suffixes and abbreviations
        # Allow shorter models (>= 2 chars) to capture models like K4, K5, X3, Q5, etc.
        potential_models = [m for m in models if m not in exclude_words and len(m) >= 2]
        
        # Try to find model near the make
        if self.data['make'] and potential_models:
            make_upper = self.data['make'].upper()
            text_upper = self.text.upper()
            make_pos = text_upper.find(make_upper)
            if make_pos != -1:
                # Look in nearby text (within 300 chars before and after for better model capture)
                start = max(0, make_pos - 150)
                end = min(len(text_upper), make_pos + 150)
                nearby_text = text_upper[start:end]
                
                # First check for common two-word models (e.g., SANTA FE, GRAND CHEROKEE)
                two_word_models = [
                    'SANTA FE', 'GRAND CHEROKEE', 'GRAND WAGONEER', 'LAND CRUISER',
                    'RANGE ROVER', 'MODEL S', 'MODEL X', 'MODEL Y', 'MODEL 3',
                    'XC40 HYB', 'XC60 HYB', 'XC90 HYB',  # Volvo hybrids
                    'TRANSIT CARGO', 'TRANSIT VAN', 'TRANSIT CONNECT'  # Ford vans
                ]
                for model in two_word_models:
                    if model in nearby_text:
                        self.data['model'] = model.title()
                        break
                
                # If no two-word model found, check single-word models
                if 'model' not in self.data or not self.data['model']:
                    for model in potential_models:
                        if model in nearby_text and model != make_upper:
                            # Check if this model appears in an email address
                            model_pos = nearby_text.find(model)
                            if model_pos != -1:
                                # Get context around the model
                                ctx_start = max(0, model_pos - 20)
                                ctx_end = min(len(nearby_text), model_pos + len(model) + 20)
                                model_context = nearby_text[ctx_start:ctx_end]
                                
                                # Skip if model is part of an email address
                                if '@' in model_context and ('.COM' in model_context or '.NET' in model_context):
                                    # Check if model is actually inside the email
                                    if re.search(model + r'[0-9]*@', model_context) or \
                                       re.search(r'@[A-Z0-9]*' + model, model_context):
                                        continue  # Skip this model, it's in an email
                            
                            self.data['model'] = model.title()
                            break
        
        if 'model' not in self.data or not self.data['model']:
            self.data['model'] = "MODEL"
        
        # HARDCODED POST-PROCESSING: Final safety check to never allow certain words as model
        # This catches any cases where these might have slipped through
        invalid_models = {'Inc', 'St', 'Llc', 'Corp', 'Ltd', 'Co', 'Inc.', 'St.', 'Llc.', 'Corp.', 'Ltd.', 'Co.', 'Hwy', 'Highway', 'Hwy.', 'Highway.'}
        if self.data.get('model') in invalid_models:
            print(f"[WARNING] Invalid model '{self.data['model']}' detected, resetting to MODEL")
            self.data['model'] = "MODEL"
        
        # HARDCODED VIN-SPECIFIC MODEL CORRECTIONS
        # VIN starting with KL4AMESL (Buick Encore GX)
        vin = self.data.get('vin', '')
        if vin.startswith('KL4AMESL'):
            self.data['model'] = "Encore Gx"
            print(f"[INFO] VIN {vin[:8]}... identified as Buick Encore GX")
        # VIN starting with KMHLM4DG (Hyundai Elantra)
        elif vin.startswith('KMHLM4DG'):
            self.data['model'] = "Elantra"
            print(f"[INFO] VIN {vin[:8]}... identified as Hyundai Elantra")
        
        # Extract addresses (more complex pattern)
        # Look for street addresses with numbers and street names, including special formats
        # Use [ \t] instead of \s to avoid matching newlines (which can combine multiple addresses)
        
        # Pattern 1: Airport/Building addresses (e.g., "JFK Airport Bldg #318", "100 Airport Blvd")
        airport_pattern = r'(?:[A-Z]{3}[ \t]+)?(?:Airport|Terminal)[ \t]+(?:Bldg|Building|Terminal)?[ \t]*(?:#[ \t]*)?[\w\d]+'
        airport_addresses = re.findall(airport_pattern, self.text, re.IGNORECASE)
        
        # Pattern 2: Standard street addresses with numbers
        address_pattern = r'\d+[ \t]+[A-Z\.][\w \t\.]+?[ \t]+(?:St|Street|Ave|Avenue|Rd|Road|Dr|Drive|Blvd|Boulevard|Ln|Lane|Way|Court|Ct|Circle|Cir|Pkwy|Parkway|Cut|Hwy|Highway|Bldg|Building)'
        street_addresses = re.findall(address_pattern, self.text, re.IGNORECASE)
        
        # Pattern 3: Street addresses without leading numbers (e.g., "Airport Blvd", "Main Street")
        # This catches street names that appear before city/state/zip
        no_number_pattern = r'(?:^|(?<=\n))([A-Z][\w \t\.]+?[ \t]+(?:St|Street|Ave|Avenue|Rd|Road|Dr|Drive|Blvd|Boulevard|Ln|Lane|Way|Court|Ct|Circle|Cir|Pkwy|Parkway|Highway|Hwy))(?=[ \t]*(?:[A-Z][a-z]+[ \t]*,[ \t]*[A-Z]{2}|\n|$))'
        no_number_addresses = re.findall(no_number_pattern, self.text, re.IGNORECASE | re.MULTILINE)
        
        # Pattern 4: Addresses without standard suffixes (e.g., "568 N. Madrid", "123 E Oak")
        # This catches numbered addresses with directionals that don't have St/Ave/Rd suffixes
        # Match: 2-4 digit number + directional (N, S, E, W, etc.) + proper noun (street name)
        # Use more specific criteria: starts with 2-4 digits, has directional, ends with capitalized word(s)
        no_suffix_pattern = r'\b(\d{2,4}\s+[NSEW]\.?\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\b'
        no_suffix_addresses = re.findall(no_suffix_pattern, self.text)
        
        # Further filter: only keep addresses that appear near a city/state/zip pattern
        # This helps ensure we're getting actual street addresses and not random number+direction+word combos
        filtered_no_suffix = []
        for addr in no_suffix_addresses:
            # Find the position of this address
            addr_pos = self.text.find(addr)
            if addr_pos != -1:
                # Look for city/state/zip within 500 characters after this address
                context = self.text[addr_pos:addr_pos + 500]
                if re.search(r'[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\s*,\s*[A-Z]{2}\s+\d{5}', context):
                    filtered_no_suffix.append(addr)
        no_suffix_addresses = filtered_no_suffix
        
        # Prioritize airport addresses (more likely to be Hertz locations), then numbered addresses with suffixes, 
        # then addresses without suffixes, then addresses without numbers
        addresses = airport_addresses + street_addresses + no_suffix_addresses + no_number_addresses
        
        # HARDCODED EXCLUSION: Filter out specific incorrect addresses
        EXCLUDED_ADDRESSES = ['221 N 36th St', '1790 S Eastwood Dr', '1875 Dekalb Ave', '2141 E Geneva St', '3200 E Jackson St']
        filtered_addresses = []
        for addr in addresses:
            # Check if this address should be excluded
            is_excluded = False
            for excluded in EXCLUDED_ADDRESSES:
                if excluded.lower() in addr.lower():
                    is_excluded = True
                    break
            if not is_excluded:
                filtered_addresses.append(addr)
        
        addresses = filtered_addresses
        
        # Also look for city, state, zip
        # Use [ ] instead of \s to avoid matching newlines (which could capture names before city)
        # Allow optional spaces before and after comma
        city_state_pattern = r'([A-Z][a-z]+(?: +[A-Z][a-z]+)*) *, *([A-Z]{2}) +(\d{5})'
        city_state_matches = re.findall(city_state_pattern, self.text)
        
        if addresses:
            # Clean up the address
            self.data['street_address'] = addresses[0].strip()
        else:
            self.data['street_address'] = "ADDRESS"
        
        if city_state_matches:
            # Use the first city/state/zip match
            city, state, zip_code = city_state_matches[0]
            self.data['city'] = city
            self.data['state'] = state
            self.data['zip'] = zip_code
            self.data['full_location'] = f"{self.data['street_address']} {city}, {state} {zip_code}"
        else:
            self.data['city'] = "CITY"
            self.data['state'] = "ST"
            self.data['zip'] = "00000"
            self.data['full_location'] = "LOCATION"
        
        # Extract phone numbers
        phone_pattern = r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}'
        phones = re.findall(phone_pattern, self.text)
        self.data['contact_phone'] = phones[0] if phones else "PHONE"
        
        # Extract order number
        # Look for numeric IDs (typically 7-8 digits)
        order_pattern = r'\b\d{7,8}\b'
        order_matches = re.findall(order_pattern, self.text)
        if order_matches:
            # Exclude VIN-like numbers and odometer readings
            self.data['order_number'] = order_matches[0]
        else:
            self.data['order_number'] = "ORDER_NUM"
        
        return self.data
    
    def generate_response_deadline(self):
        """Generate response deadline timestamp - 1100 on next business day."""
        next_day = datetime.now() + timedelta(days=1)
        
        # Check if next day is a weekend, if so, move to Monday
        # weekday(): Monday=0, Tuesday=1, ..., Saturday=5, Sunday=6
        while next_day.weekday() >= 5:  # 5=Saturday, 6=Sunday
            next_day = next_day + timedelta(days=1)
        
        return next_day.strftime("1100 on %m/%d")
    
    def create_email_subject(self):
        """Create email subject line."""
        vin = self.data.get('vin', 'VIN')
        # You can customize CONDITION/LOCATION based on data or make it editable
        return f"HERTZ LOGISTICS CONFIRMATION - VIN # ({vin}) - CONDITION/LOCATION - RESPONSE NEEDED BY {self.generate_response_deadline()}"
    
    def _mileage_verification_html(self):
        """Return mileage verification verbiage for vehicles model year 2023 or older."""
        year_str = self.data.get('year', '')
        try:
            year_int = int(year_str)
        except (ValueError, TypeError):
            return ""
        if year_int <= 2023:
            return (
                '<p style="margin-top: 12px;"><b>*** MILEAGE VERIFICATION REQUIRED ***</b><br>'
                'This unit is a 2023 or older model year vehicle. '
                'Please verify and confirm the current mileage/odometer reading on the unit.</p>'
            )
        return ""

    def create_email_body(self):
        """Create email body text in HTML format with bold formatting."""
        year = self.data.get('year', 'YEAR')
        make = self.data.get('make', 'MAKE')
        model = self.data.get('model', 'MODEL')
        vin = self.data.get('vin', 'VIN')
        location = self.data.get('full_location', 'LOCATION')
        
        # Create HTML body with bold VIN, vehicle info, and address
        html_body = f"""<html>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 12pt;">
<p>Hello,</p>
 
<p>Please confirm this unit is on site at <b>{location}</b> – Hertz is transporting to the dealer</p>
 
<p><b>{vin} – {year} {make} {model}</b></p>
 
<p>Please answer the following questions regarding the status of the unit so we can ensure its in good condition to be sold.</p>
 
<ol>
<li>Is there any damage on the body of this vehicle? Y or N</li>
<li>Is there any glass damage (chips or cracks)? Y or N</li>
<li>Is there any paint mismatch (likely from previous repair) Y or N</li>
<li>Are all the tires in good condition (no flats or low tires)? Y or N</li>
<li>Is the interior clean and debris free? Y or N</li>
</ol>
{self._mileage_verification_html()}
 
<p>Best regards,<br>
Hertz Transportation Team</p>
</body>
</html>"""
        return html_body
    
    def create_outlook_draft(self):
        """Create an Outlook email draft with the extracted information."""
        try:
            # Initialize Outlook
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # 0 = MailItem
            
            # Set recipients (Hertz rep emails)
            if self.data.get('hertz_rep_emails'):
                mail.To = "; ".join(self.data['hertz_rep_emails'])
            
            # Set CC (always CC joshua.blankenship@hertz.com)
            # Check if config file exists and has CC_RECIPIENTS
            cc_recipients = "joshua.blankenship@hertz.com"  # Default
            try:
                import email_config
                if hasattr(email_config, 'CC_RECIPIENTS'):
                    cc_recipients = email_config.CC_RECIPIENTS
            except ImportError:
                pass  # Use default
            
            mail.CC = cc_recipients
            
            # Set subject
            mail.Subject = self.create_email_subject()
            
            # Set body (using HTML for formatting)
            mail.HTMLBody = self.create_email_body()
            
            # Attach the PDF file
            import os
            pdf_full_path = os.path.abspath(self.pdf_path)
            if os.path.exists(pdf_full_path):
                mail.Attachments.Add(pdf_full_path)
                print(f"[OK] PDF attached: {os.path.basename(pdf_full_path)}")
            else:
                print(f"[WARNING] PDF file not found for attachment: {pdf_full_path}")
            
            # Display the draft (don't send automatically)
            mail.Display()
            
            print("[OK] Outlook email draft created successfully!")
            return True
            
        except Exception as e:
            print(f"[ERROR] Error creating Outlook draft: {e}")
            print("  Make sure Microsoft Outlook is installed and configured.")
            return False
    
    def print_summary(self):
        """Print extracted data summary."""
        print("\n" + "="*80)
        print("EXTRACTED DATA SUMMARY")
        print("="*80)
        print(f"\nPDF File: {self.pdf_path}")
        print(f"\nHertz Representative Emails (TO Recipients):")
        if self.data.get('hertz_rep_emails'):
            for email in self.data.get('hertz_rep_emails', []):
                print(f"  - {email}")
        else:
            print("  - None found")
        
        # Get CC recipient from config
        cc_recipients = "joshua.blankenship@hertz.com"  # Default
        try:
            import email_config
            if hasattr(email_config, 'CC_RECIPIENTS'):
                cc_recipients = email_config.CC_RECIPIENTS
        except ImportError:
            pass
        
        print(f"\nCC Recipients:")
        print(f"  - {cc_recipients} [ALWAYS CC'd]")
        
        # Show excluded emails
        excluded_emails = self.data.get('excluded_emails', [])
        if excluded_emails:
            print(f"\nExcluded Emails (Purchaser):")
            for email in excluded_emails:
                print(f"  - {email} [EXCLUDED - First email/Purchaser]")
        print(f"\nVehicle Information:")
        print(f"  VIN:   {self.data.get('vin', 'N/A')}")
        print(f"  Year:  {self.data.get('year', 'N/A')}")
        print(f"  Make:  {self.data.get('make', 'N/A')}")
        print(f"  Model: {self.data.get('model', 'N/A')}")
        print(f"\nPickup Location:")
        print(f"  {self.data.get('full_location', 'N/A')}")
        print(f"\nContact Phone:")
        print(f"  {self.data.get('contact_phone', 'N/A')}")
        print(f"\nOrder Number:")
        print(f"  {self.data.get('order_number', 'N/A')}")
        print("\n" + "="*80 + "\n")


def process_pdf(pdf_path, auto_create_email=False):
    """Main function to process a PDF and create email draft."""
    print(f"\nProcessing PDF: {pdf_path}")
    print("-" * 80)
    
    processor = HertzPDFProcessor(pdf_path)
    processor.extract_text()
    processor.parse_data()
    processor.print_summary()
    
    # Ask user to confirm before creating draft (unless auto mode)
    if auto_create_email:
        print("\nAuto-creating Outlook email draft...")
        processor.create_outlook_draft()
    else:
        try:
            response = input("Create Outlook email draft? (Y/n): ").strip().lower()
            if response in ['y', 'yes', '']:
                processor.create_outlook_draft()
            else:
                print("Email draft creation cancelled.")
        except EOFError:
            print("\nNon-interactive mode detected. Use --auto flag to create draft automatically.")
    
    return processor


def main():
    """Main entry point."""
    print("="*80)
    print("HERTZ PDF SCRUBBER & EMAIL GENERATOR")
    print("="*80)
    
    # Check for command-line arguments
    auto_mode = '--auto' in sys.argv or '-a' in sys.argv
    
    # Check if a PDF file was specified as argument
    pdf_arg = None
    for arg in sys.argv[1:]:
        if arg.endswith('.pdf') or Path(arg).suffix.lower() == '.pdf':
            pdf_arg = arg
            break
    
    if pdf_arg:
        pdf_path = pdf_arg
        print(f"\nUsing specified file: {pdf_path}")
    else:
        # Look for PDF files in current directory
        pdf_files = list(Path('.').glob('*.pdf'))
        
        if not pdf_files:
            print("\n✗ No PDF files found in current directory.")
            print("\nUsage: python hertz_email_generator.py [PDF_FILE] [--auto]")
            print("  --auto, -a    Automatically create email draft without confirmation")
            return
        else:
            print(f"\nFound {len(pdf_files)} PDF file(s):")
            for idx, pdf in enumerate(pdf_files, 1):
                print(f"  {idx}. {pdf.name}")
            
            if len(pdf_files) == 1:
                pdf_path = str(pdf_files[0])
                print(f"\nUsing: {pdf_path}")
            else:
                try:
                    choice = input(f"\nSelect file (1-{len(pdf_files)}): ").strip()
                    idx = int(choice) - 1
                    pdf_path = str(pdf_files[idx])
                except (ValueError, IndexError, EOFError):
                    print("Using first file.")
                    pdf_path = str(pdf_files[0])
    
    # Process the PDF
    process_pdf(pdf_path, auto_create_email=auto_mode)


if __name__ == "__main__":
    main()

