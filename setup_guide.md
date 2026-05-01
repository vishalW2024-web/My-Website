# 3D QR Menu — Complete Setup Guide

## ✅ What's Working Right Now

| Feature | Status | Details |
|---------|--------|---------|
| Portfolio (Google Sheet) | ✅ Live | Dezzert Dudess shows with gold QR code |
| Google Forms Backend | ✅ Connected | Form submissions go to your Google Form |
| WhatsApp Integration | ✅ Active | +91 8459987710 |
| 3D Intro Splash | ✅ Working | Floating QR codes |
| Cursor Follower | ✅ Working | Changes per section |
| Image Library Page | ✅ Ready | Opens on click, will show categories when Sheet2 is set up |

---

## 📋 Google Form — Receiving Entries

**Your form is CONNECTED and WORKING!** We tested it and the submission succeeded.

### To verify entries are received:

1. Go to your Google Form: [Open Form](https://docs.google.com/forms/d/e/1FAIpQLScGgh3MmzImq8FR4zLxraT0XB9nzfj6oEYmMtIavCZRAPFQFg/viewform)
2. Click the **"Responses"** tab at the top
3. You should see entries there (including our test submission)

### Entry IDs mapped:
| Website Field | Google Form Entry ID |
|---------------|---------------------|
| Full Name | `entry.2077733847` |
| Restaurant Name | `entry.163609028` |
| Business Email | `entry.1621135527` |
| Phone Number | `entry.1968309570` |
| Menu Link | `entry.905678312` |
| Message | `entry.1418583208` |

---

## 📧 FREE Auto-Reply Email Setup (Google Apps Script)

> [!IMPORTANT]
> This is **100% free** using Google Apps Script. When someone submits the form, they'll automatically receive a professional email from your Gmail.

### Step-by-Step Instructions:

1. **Open your Google Form** in edit mode
2. Click the **⋮ (three dots)** menu → **"Script editor"**
3. Delete everything in the editor and paste this script:

```javascript
function onFormSubmit(e) {
  var responses = e.namedValues;
  
  var recipientEmail = responses['Business Email Address'][0];
  var name = responses['Your Full Name (Contact Person)'][0];
  var restaurant = responses['Restaurant Name'][0];
  
  var subject = '🎉 Welcome to 3D QR Menu — We Received Your Inquiry!';
  
  var htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#0a0e1a;color:#e0e0e0;padding:40px;border-radius:12px">
      <div style="text-align:center;margin-bottom:30px">
        <h1 style="color:#d4af37;font-size:28px;margin:0">3D QR Menu</h1>
        <p style="color:#888;font-size:14px">Premium Digital Menus</p>
      </div>
      
      <p style="font-size:16px">Hi <strong>${name}</strong>,</p>
      
      <p>Thank you for your interest in a QR digital menu for <strong>${restaurant}</strong>! 🎊</p>
      
      <p>We've received your inquiry and our team will review your requirements within <strong>24 hours</strong>.</p>
      
      <div style="background:#111827;border:1px solid #d4af3733;border-radius:8px;padding:20px;margin:20px 0">
        <h3 style="color:#d4af37;margin-top:0">What happens next?</h3>
        <ol style="line-height:2">
          <li>✅ Our design team reviews your requirements</li>
          <li>📱 We'll reach out via WhatsApp for a quick discussion</li>
          <li>🎨 We create a custom design mockup for your menu</li>
          <li>📲 Your QR menu goes live within 48 hours</li>
        </ol>
      </div>
      
      <div style="background:#111827;border-radius:8px;padding:20px;margin:20px 0">
        <h3 style="color:#d4af37;margin-top:0">What's Included (FREE):</h3>
        <ul style="line-height:2">
          <li>🆓 Free hosting forever</li>
          <li>🆓 Free QR code generation</li>
          <li>🆓 Free professional image enhancement</li>
          <li>🆓 Free updates anytime</li>
          <li>📱 Mobile-optimized luxury design</li>
        </ul>
      </div>
      
      <p>Want to discuss right away? Chat with us on WhatsApp:</p>
      
      <div style="text-align:center;margin:24px 0">
        <a href="https://wa.me/918459987710?text=Hi!%20I%20just%20submitted%20a%20form%20for%20${encodeURIComponent(restaurant)}" 
           style="background:#25D366;color:white;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:bold;display:inline-block">
          💬 Chat on WhatsApp
        </a>
      </div>
      
      <hr style="border:none;border-top:1px solid #333;margin:30px 0">
      
      <p style="color:#888;font-size:13px;text-align:center">
        3D QR Menu — Premium Digital Menus for Restaurants<br>
        📧 vishalwork2024@gmail.com | 📱 +91 8459987710
      </p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: htmlBody,
    name: '3D QR Menu',
    replyTo: 'vishalwork2024@gmail.com'
  });
  
  // Also notify yourself
  MailApp.sendEmail({
    to: 'vishalwork2024@gmail.com',
    subject: '🔔 New QR Menu Inquiry — ' + restaurant,
    htmlBody: '<h2>New Lead!</h2><p><strong>Name:</strong> ' + name + '</p><p><strong>Restaurant:</strong> ' + restaurant + '</p><p><strong>Email:</strong> ' + recipientEmail + '</p><p><strong>Message:</strong> ' + (responses['Your Message/Specific Inquiry'] || [''])[0] + '</p>',
    name: '3D QR Menu Leads'
  });
}
```

4. **Save** the script (Ctrl+S)
5. Click **"Triggers"** (clock icon on left sidebar) → **"+ Add Trigger"**
6. Set:
   - Function: `onFormSubmit`
   - Event source: `From form`
   - Event type: `On form submit`
7. Click **Save** → Authorize with your Google account
8. **Done!** Every form submission now sends a beautiful email to the customer AND a notification to you.

> [!TIP]
> Google Apps Script lets you send up to **100 emails/day** for free. More than enough for lead generation!

---

## 🖼️ Image Library — Setting Up Sheet2

To show category-wise before/after images in the Image Library page:

### Step 1: Add a new tab to your Google Sheet

1. Open your Google Sheet (the one with "Dezzert Dudess")
2. Click the **"+"** button at the bottom to add a new sheet tab
3. Name it **"Image Library"**

### Step 2: Set up the columns

| Category | Item Name | Raw Image Link | Enhanced Image Link |
|----------|-----------|----------------|---------------------|
| Desserts | Chocolate Cake | (Google Drive link to raw photo) | (Google Drive link to enhanced photo) |
| Main Course | Butter Chicken | (Drive link) | (Drive link) |
| Beverages | Cold Coffee | (Drive link) | (Drive link) |

### Step 3: Publish Sheet2

1. Go to **File → Share → Publish to web**
2. Select **"Image Library"** tab (Sheet2)
3. Select **CSV** format
4. Click **Publish**
5. Copy the URL
6. **Send me that URL** and I'll add it to the website config

> [!NOTE]
> For the image links, upload images to Google Drive, right-click → "Share" → "Anyone with the link", then paste the share link. The website will automatically convert them for display.

---

## 📊 Portfolio Sheet — Adding More Entries

Your Sheet1 structure:

| Name of the Outlet | QR Image Link |
|---------------------|---------------|
| Dezzert Dudess | (your link) |
| New Restaurant | (menu URL or QR image link) |

- Add new rows and they appear on the website automatically
- The website generates a **gold-themed scannable QR code** from whatever link you provide
- If you have 3+ entries, a "**See All Menus**" button appears automatically
