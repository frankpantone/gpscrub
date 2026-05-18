# PDF Structure Analysis Report
## File: GR4031029_KNDPU3AF3P7143419.pdf

---

## Document Overview
- **Document Type:** Gate Release / Vehicle Pickup Form
- **Pages:** 1
- **Total Characters:** 1,459
- **Email Addresses Found:** 3
- **Phone Numbers Found:** 4

---

## Email Addresses to Scrub

| Email Address | Context/Role | Line Number |
|--------------|--------------|-------------|
| `Khadijah.Goudeau@hertz.com` | Contact Email | 43 |
| `BSwift@hertz.com` | Contact Email | 45 |
| `konovak@hertz.com` | Contact Email | 37 |

---

## Document Structure Map

### SECTION 1: Pickup Location Information
**Field Labels (Placeholders):**
- Pickup Location Gate Release Code
- Contact
- Phone
- Address
- Email
- Hours of Operation

**Actual Values:**
- Name: `Name` (line 28)
- Phone: `(317) 775-8055` (lines 29-30)
- Hours: `M-F 8am - 3pm` (line 31)

---

### SECTION 2: Order Information
**Field Labels (Placeholders):**
- Order Number
- Instruction

**Actual Values:**
- Order Number: `4031029` (line 35)
- Instruction: `Brian Swift/Khadijah Goudeau` (line 32)

---

### SECTION 3: Vehicle Information
**Field Labels (Placeholders):**
- Owning Area Number
- Unit Number
- Serial
- Odometer
- Model Year
- Color
- Make
- Model
- Body Type / Pkg
- License #
- Noted Damage

**Actual Values:**
- Owning Area: `RHP AP INDIANAPOLIS` (line 55)
- Unit Number: `2836` (line 56)
- Serial: `KNDPU3AF3P7143419` (line 64)
- Odometer: `66855` (line 51)
- Model Year: `2023` (line 58)
- Color: `GRAY` (line 33)
- Make: `KIA` (line 47)
- Model: `SPORTAGE` (line 46)
- Body Type: `FWD` (line 44)
- License: `QNFB46` (line 62)
- VL Code: `VL3NRTBB` (line 57)

---

### SECTION 4: Purchaser Information
**Field Labels (Placeholders):**
- Representative
- Manager
- Office #
- Cell #
- Fax #
- Name
- Contact
- Address
- Phone
- Email

**Actual Values:**
- Name: `Kori Novak` (lines 48, 59)
- Contact: (appears to be `Mike Bloom` at line 34)
- Address: `221 N 36th St` (line 39), `Quincy, IL 62301` (line 60)
- Phone: `4142179780` (lines 36, 61) - formatted as 414-217-9780
- Email: **`konovak@hertz.com`** (line 37) ⚠️ EMAIL TO SCRUB

---

### SECTION 5: Hertz Rep Information
**Field Labels (Placeholders):**
- Hertz Rep Information

**Actual Values:**
- Contact: `Kori Novak` (line 59)
- Phone: `8662225012` (line 49) - formatted as 866-222-5012
- Email: **`Khadijah.Goudeau@hertz.com`** (line 43) ⚠️ EMAIL TO SCRUB
- Email: **`BSwift@hertz.com`** (line 45) ⚠️ EMAIL TO SCRUB
- Address: `2825 W. Perimeter Rd` (line 38), `Indianapolis, IN 46241` (line 41)

---

### SECTION 6: Driver/Transporter Information
**Field Labels (Placeholders):**
- Transport Company / Dealership
- Driver / Transporter Printed Name
- Driver / Transporter License Number
- Driver / Transporter Signature
- Hertz Authorized Signature
- Date Picked Up

**Actual Values:**
- Transport Company: `Kunes Country Honda of Quincy Inc` (line 42)
- Date: `2025-10-28` (line 40)
- Other values: `2151439` (line 50) - possibly license or ID number
- Phone: `2172287000` (line 65) - formatted as 217-228-7000

---

## Footer/Legal Text
**Field Labels (Placeholders):**
- Legal disclaimer text

**Actual Values:**
```
Transporter and Purchaser agree to assume responsibility for any risk,
loss or damage to vehicle(s) at such time as the vehicle(s) are received
and accepted by Transporter or Purchaser.
```

**Instructions:**
```
TRANSPORTER MUST PRESENT THIS GATE RELEASE AT TIME OF PICK UP. 
VEHICLE WILL NOT BE RELEASED WITHOUT IT

There is no Hertz buildings or employees on this lot. 
Please call us at (317) 775-8055 and we will come meet you on this lot
```

---

## Format Analysis

### Document Layout Pattern
The PDF follows a **label-above-value** or **mixed inline-value** format where:
1. Field labels are printed as structural elements (placeholders)
2. Actual data values are filled in adjacent to or below the labels
3. Some sections have nested indentation (indicated by spacing)

### Field Identification Logic

**PLACEHOLDERS (Field Labels):**
- Typically appear as standalone descriptive text
- Often end with colons or are part of section headers
- Examples: "Email", "Phone", "Order Number", "Serial", etc.

**VALUES (Actual Data):**
- Concrete data that would change between forms
- Personal information (names, emails, phone numbers)
- Vehicle details (VIN, odometer, color)
- Dates, addresses, company names
- Numeric identifiers

---

## Email Scrubbing Strategy

### Recommended Approach
Replace email addresses with a placeholder pattern that maintains document format:

**Option 1: Generic Replacement**
```
[EMAIL_REMOVED]
```

**Option 2: Role-Based Replacement**
```
[CONTACT_EMAIL]
[HERTZ_REP_EMAIL]
[PURCHASER_EMAIL]
```

**Option 3: Format-Preserving Replacement**
```
[redacted]@hertz.com
contact@[redacted].com
```

### Emails to Replace

1. **Line 37:** `konovak@hertz.com`
   - Context: Purchaser contact email
   - Associated with: Kori Novak

2. **Line 43:** `Khadijah.Goudeau@hertz.com`
   - Context: Hertz representative email
   - Associated with: Instruction contact (Brian Swift/Khadijah Goudeau)

3. **Line 45:** `BSwift@hertz.com`
   - Context: Hertz representative email
   - Associated with: Instruction contact (Brian Swift/Khadijah Goudeau)

---

## Summary

This gate release form contains:
- ✅ **Structured field labels** that serve as placeholders/templates
- ✅ **Filled values** representing actual transaction data
- ⚠️ **3 email addresses** that need to be scrubbed
- ⚠️ **4 phone numbers** (may also need scrubbing depending on requirements)
- ℹ️ **Personal names** and **addresses** (consider if these need scrubbing too)

The document has a clear hierarchical structure with sections for:
1. Location/Contact Info
2. Order Details
3. Vehicle Details
4. Purchaser Info
5. Hertz Representative Info
6. Driver/Transporter Info
7. Legal/Instructions

**Next Steps:** Create a scrubbing script that:
1. Extracts text from PDF
2. Identifies and replaces email addresses
3. Optionally replaces other PII (phone, names, addresses)
4. Generates a new "scrubbed" PDF maintaining the original layout




