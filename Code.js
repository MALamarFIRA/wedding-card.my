// --- Code.gs Backend ---

function doGet() {
  // Serves the front-end Index.html template
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Wedding Invitation: Akmal & Musfirah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 🟢 GLOBAL CONFIGURATION (Safe for GitHub - reads securely from your vault)
const SHEET_ID = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
const TARGET_SHEET_NAME = "RSVP"; // Ensure this matches the exact tab name

function saveRsvpData(formData) {
  try {
    // Total reliability check: Open by ID using our secure global variable
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
    // Safety check if sheet name is misspelled
    if (!sheet) {
      throw new Error(`Error: Sheet tab named "${TARGET_SHEET_NAME}" was not found.`);
    }
    
    // 2. Efficient Duplicate Validation Match Engine
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      // getValues returns a 2D array [ ['Name1'], ['Name2'], etc. ]
      const existingNamesData = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); 
      
      // Clean up incoming name for comparison (removes spaces, ignores case)
      const incomingName = formData.name.trim().toLowerCase();
      
      // Loop through the 2D array to check for duplicates
      for (var i = 0; i < existingNamesData.length; i++) {
        const currentSheetName = existingNamesData[i][0].toString().trim().toLowerCase();
        
        if (currentSheetName === incomingName) {
          return "Duplicate"; // Trigger the front-end duplicate alert pop-up
        }
      }
    }
    
    // 3. Save RSVP row if no duplicate is found
    // Final Data Layout: Timestamp, Full Name, Status, Pax, Email, Wishes
    sheet.appendRow([
      new Date(), 
      formData.name.trim(), 
      formData.status, 
      formData.pax, 
      formData.email, 
      formData.wishes
    ]);
    
    return "Success"; // Trigger the front-end 'RSVP Recorded' alert
    
  } catch(error) {
    console.error("Backend RSVP Save Error:", error);
    return "Error: " + error.toString();
  }
}

function getExistingWishes() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) return []; // Safety return if sheet is missing
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // Return empty if there's no data yet

    const data = sheet.getDataRange().getValues();
    const wishesList = [];
    
    // Loops through data rows starting after the header row (i = 1)
    for (var i = 1; i < data.length; i++) {
      var name = data[i][1]; // Column B (Index 1)
      var wish = data[i][5]; // Column F (Index 5)
      
      // Only pull entries that actually wrote a wish
      if (wish && wish.toString().trim() !== "") {
        wishesList.push({
          name: name,
          wishes: wish
        });
      }
    }
    
    // Return them reversed so the newest submissions show up first
    return wishesList.reverse(); 
    
  } catch(error) {
    console.error("Backend Wish Fetch Error:", error);
    return [];
  }
}

function personalassistant(userPrompt) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  
  if (!API_KEY) {
    return "Backend Configuration Error: GEMINI_API_KEY property is missing.";
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
  
  const systemInstruction = `
    You are an AI Travel & Logistics Concierge for Akmal & Musfirah's wedding. Use the official information below to answer the user's questions perfectly.
    
    WEDDING DETAILS:
    - Date: 1 May, 2027 (Saturday)
    - Venue: Grandball Room, Eastwood Valley Golf & Country Club, Miri, Sarawak, Malaysia.
    - Venue Google Maps Link: https://maps.app.goo.gl/kx7TFCVeTM857Pax5
    - Venue Waze Maps Link: https://ul.waze.com/ul?venue_id=74711084.747241909.2804519&overview=yes&utm_campaign=default&utm_source=waze_website&utm_medium=lm_share_location
    
    FLIGHT BOOKING LINKS:
    Direct flights route from Kuala Lumpur (KUL) to Miri International Airport (MYY). Instruct guests to use these official links to book:
    - AirAsia: https://www.airasia.com
    - Malaysia Airlines: https://www.malaysiaairlines.com
    
    HOTEL OPTIONS, DISTANCES & GRAB FARES:
    Provide these options when guests ask about accommodations, hotels, or where to stay near the event venue:
    1. Eastwood Valley Golf & Country Club (On-site venue rooms/chalets)
       - Distance: 0 km (Inside the wedding venue area)
       - Grab Fare: N/A (Walkable to the venue)
       - Booking links: Agoda: https://www.agoda.com/eastwood-valley-golf-country-club/hotel/miri-my.html?countryId=198&finalPriceView=1&isShowMobileAppPrice=false&cid=-999&numberOfBedrooms=&familyMode=false&adults=1&children=0&rooms=1&maxRooms=0&checkIn=2027-04-30&isCalendarCallout=false&childAges=&numberOfGuest=0&missingChildAges=false&travellerType=0&showReviewSubmissionEntry=false&currencyCode=MYR&isFreeOccSearch=false&flightSearchCriteria=%5Bobject+Object%5D&tspTypes=16&los=2&searchrequestid=253f69ea-dbba-4f29-a9f2-e912e2c0402c&ds=wkEf3JY%2Fgn%2B40%2Baz | Booking.com: https://www.booking.com/hotel/my/eastwood-valley-golf-country-club.html?aid=304142&label=gog235jc-10CAEoggI46AdIM1gDaKEBiAEBmAEzuAEXyAEM2AED6AEB-AEBiAIBqAIBuALP8_rRBsACAdICJDhlN2IyOTNiLTg5YTQtNDZiZi04YTc0LTNkZjE1MWMyOGMxZdgCAeACAQ&sid=3b8237052f56173a4d8f79efe205c7f1&all_sr_blocks=23817204_392628049_0_1_0&checkin=2027-04-30&checkout=2027-05-02&dest_id=238172&dest_type=hotel&dist=0&group_adults=1&group_children=0&hapos=1&highlighted_blocks=23817204_392628049_0_1_0&hpos=1&matching_block_id=23817204_392628049_0_1_0&no_rooms=1&req_adults=1&req_children=0&room1=A&sb_price_type=total&sr_order=popularity&sr_pri_blocks=23817204_392628049_0_1_0__30400&srepoch=1782496166&srpvid=d14c7d51c519014e&type=total&ucfs=1&
       
    2. Pullman Miri Waterfront (Luxury 5-star option)
       - Distance: ~7.5 km from venue (~12 mins drive)
       - Estimated Grab Fare: RM 11.00 - RM 16.00
       - Booking links: Agoda: https://www.agoda.com/pullman-miri-waterfront/hotel/miri-my.html?countryId=198&finalPriceView=1&isShowMobileAppPrice=false&cid=-999&numberOfBedrooms=&familyMode=false&adults=1&children=0&rooms=1&maxRooms=0&checkIn=2027-04-30&isCalendarCallout=false&childAges=&numberOfGuest=0&missingChildAges=false&travellerType=0&showReviewSubmissionEntry=false&currencyCode=MYR&isFreeOccSearch=false&flightSearchCriteria=%5Bobject+Object%5D&tspTypes=16&los=2&searchrequestid=560edd80-4e53-49c2-a0e1-3b59ca2a2cd1&ds=wkEf3JY%2Fgn%2B40%2Baz | Booking.com: https://www.booking.com/hotel/my/pullman-miri-waterfront.html?aid=304142&label=gog235jc-10CAEoggI46AdIM1gDaKEBiAEBmAEzuAEXyAEM2AED6AEB-AEBiAIBqAIBuALP8_rRBsACAdICJDhlN2IyOTNiLTg5YTQtNDZiZi04YTc0LTNkZjE1MWMyOGMxZdgCAeACAQ&sid=3b8237052f56173a4d8f79efe205c7f1&all_sr_blocks=170383303_94612381_2_2_0_729336&checkin=2027-04-30&checkout=2027-05-02&dest_id=1703833&dest_type=hotel&dist=0&group_adults=1&group_children=0&hapos=1&highlighted_blocks=170383303_94612381_2_2_0_729336&hpos=1&matching_block_id=170383303_94612381_2_2_0_729336&no_rooms=1&req_adults=1&req_children=0&room1=A&sb_price_type=total&sr_order=popularity&sr_pri_blocks=170383303_94612381_2_2_0_729336_48960&srepoch=1782496344&srpvid=7a697d959a550fec&type=total&ucfs=1&
       
    3. Meritz Hotel Miri (City center, attached to Bintang Megamall)
       - Distance: ~6.2 km from venue (~10 mins drive)
       - Estimated Grab Fare: RM 9.00 - RM 14.00
       - Booking links: Agoda: https://www.agoda.com/meritz-hotel/hotel/miri-my.html?countryId=198&finalPriceView=1&isShowMobileAppPrice=false&cid=-999&numberOfBedrooms=&familyMode=false&adults=1&children=0&rooms=1&maxRooms=0&checkIn=2027-04-30&isCalendarCallout=false&childAges=&numberOfGuest=0&missingChildAges=false&travellerType=0&showReviewSubmissionEntry=false&currencyCode=MYR&isFreeOccSearch=false&flightSearchCriteria=%5Bobject+Object%5D&tspTypes=16&los=2&searchrequestid=8f57d4f5-6801-4854-9154-42c0b22262a7&ds=wkEf3JY%2Fgn%2B40%2Baz | Booking.com: https://www.booking.com/hotel/my/meritz.html?aid=304142&label=gog235jc-10CAEoggI46AdIM1gDaKEBiAEBmAEzuAEXyAEM2AED6AEB-AEBiAIBqAIBuALP8_rRBsACAdICJDhlN2IyOTNiLTg5YTQtNDZiZi04YTc0LTNkZjE1MWMyOGMxZdgCAeACAQ&sid=3b8237052f56173a4d8f79efe205c7f1&all_sr_blocks=39426524_0_1_1_0&checkin=2027-04-30&checkout=2027-05-02&dest_id=394265&dest_type=hotel&dist=0&group_adults=1&group_children=0&hapos=1&highlighted_blocks=39426524_0_1_1_0&hpos=1&matching_block_id=39426524_0_1_1_0&no_rooms=1&req_adults=1&req_children=0&room1=A&sb_price_type=total&sr_order=popularity&sr_pri_blocks=39426524_0_1_1_0__51593&srepoch=1782496482&srpvid=ae547de3e6480219&type=total&ucfs=1&
       
    4. Imperial Hotel Miri / Imperial Palace Hotel (Comfortable business hotels)
       - Distance: ~6.5 km from venue (~11 mins drive)
       - Estimated Grab Fare: RM 10.00 - RM 15.00
       - Booking links: Agoda: https://www.agoda.com/imperial-palace-hotel/hotel/miri-my.html?countryId=198&finalPriceView=1&isShowMobileAppPrice=false&cid=-999&numberOfBedrooms=&familyMode=false&adults=1&children=0&rooms=1&maxRooms=0&checkIn=2027-04-30&isCalendarCallout=false&childAges=&numberOfGuest=0&missingChildAges=false&travellerType=0&showReviewSubmissionEntry=false&currencyCode=MYR&isFreeOccSearch=false&flightSearchCriteria=%5Bobject+Object%5D&los=2&searchrequestid=3bd2cab9-bc93-4c85-8edd-eabc5b4ac147&ds=wkEf3JY%2Fgn%2B40%2Baz | Booking.com: https://www.booking.com/hotel/my/imperial-miri.html?aid=304142&label=gog235jc-10CAEoggI46AdIM1gDaKEBiAEBmAEzuAEXyAEM2AED6AEB-AEBiAIBqAIBuALP8_rRBsACAdICJDhlN2IyOTNiLTg5YTQtNDZiZi04YTc0LTNkZjE1MWMyOGMxZdgCAeACAQ&sid=3b8237052f56173a4d8f79efe205c7f1&checkin=2027-04-30&checkout=2027-05-02&dest_id=1131782&dest_type=hotel&dist=0&group_adults=1&group_children=0&hapos=1&hpos=1&no_rooms=1&req_adults=1&req_children=0&room1=A&sb_price_type=total&soh=1&sr_order=popularity&srepoch=1782496507&srpvid=5ea67dfbda3b02e1&type=total&ucfs=1&#no_availability_msg
       
    5. Kingwood Boutique Hotel Miri (Trendy, mid-range budget option)
       - Distance: ~7.0 km from venue (~12 mins drive)
       - Estimated Grab Fare: RM 10.00 - RM 15.00
       - Booking links: Agoda: https://www.agoda.com/kingwood-boutique-hotel/hotel/miri-my.html?countryId=198&finalPriceView=1&isShowMobileAppPrice=false&cid=-999&numberOfBedrooms=&familyMode=false&adults=1&children=0&rooms=1&maxRooms=0&checkIn=2027-04-30&isCalendarCallout=false&childAges=&numberOfGuest=0&missingChildAges=false&travellerType=0&showReviewSubmissionEntry=false&currencyCode=MYR&isFreeOccSearch=false&flightSearchCriteria=%5Bobject+Object%5D&tspTypes=16&los=2&searchrequestid=0c3ca728-8f8e-4579-af27-54bfce5a426b&ds=wkEf3JY%2Fgn%2B40%2Baz | Booking.com: https://www.booking.com/hotel/my/po-kingwood-boutique-hotel-the-best-rated-area-in-miri.html?aid=304142&label=gog235jc-10CAEoggI46AdIM1gDaKEBiAEBmAEzuAEXyAEM2AED6AEB-AEBiAIBqAIBuALP8_rRBsACAdICJDhlN2IyOTNiLTg5YTQtNDZiZi04YTc0LTNkZjE1MWMyOGMxZdgCAeACAQ&sid=3b8237052f56173a4d8f79efe205c7f1&all_sr_blocks=0_0_1_0_0&checkin=2027-04-30&checkout=2027-05-02&dest_id=13095006&dest_type=hotel&dist=0&group_adults=1&group_children=0&hapos=1&highlighted_blocks=0_0_1_0_0&hpos=1&matching_block_id=0_0_1_0_0&no_rooms=1&req_adults=1&req_children=0&room1=A&sb_price_type=total&sr_order=popularity&sr_pri_blocks=0_0_1_0_0__43967&srepoch=1782496641&srpvid=fa207e3fd04d0068&type=total&ucfs=1&

    6. Amigo Hotel Miri (Highly rated budget/backpacker option in the city)
       - Distance: ~7.8 km from venue (~13 mins drive)
       - Estimated Grab Fare: RM 11.00 - RM 17.00
       - Booking links: Agoda: https://www.agoda.com/amigo-hotel/hotel/miri-my.html?countryId=198&finalPriceView=1&isShowMobileAppPrice=false&cid=-999&numberOfBedrooms=&familyMode=false&adults=1&children=0&rooms=1&maxRooms=0&checkIn=2027-04-30&isCalendarCallout=false&childAges=&numberOfGuest=0&missingChildAges=false&travellerType=0&showReviewSubmissionEntry=false&currencyCode=MYR&isFreeOccSearch=false&flightSearchCriteria=%5Bobject+Object%5D&tspTypes=16&los=2&searchrequestid=6c331d4a-14ad-4cfc-b136-35a55f631bb1&ds=wkEf3JY%2Fgn%2B40%2Baz | Booking.com: booking.com/hotel/my/amigo-miri.html?label=gog235jc-10CAEoggI46AdIM1gDaKEBiAEBmAEzuAEXyAEM2AED6AEB-AEBiAIBqAIBuALP8_rRBsACAdICJDhlN2IyOTNiLTg5YTQtNDZiZi04YTc0LTNkZjE1MWMyOGMxZdgCAeACAQ&aid=304142&ucfs=1&arphpl=1&checkin=2027-04-30&checkout=2027-05-02&dest_id=4281399&dest_type=hotel&group_adults=1&req_adults=1&no_rooms=1&group_children=0&req_children=0&hpos=1&hapos=1&sr_order=popularity&srpvid=599f7e609e590bde&srepoch=1782496745&all_sr_blocks=428139902_0_1_0_0&highlighted_blocks=428139902_0_1_0_0&matching_block_id=428139902_0_1_0_0&sr_pri_blocks=428139902_0_1_0_0__29613&from=searchresults

    7. Mercure Miri City Centre (Brand new, stylish city center hotel)
       - Distance: ~6.8 km from venue (~11 mins drive)
       - Estimated Grab Fare: RM 10.00 - RM 15.00
       - Booking links: Agoda: https://www.agoda.com/mercure-miri-city-centre/hotel/miri-my.html?countryId=198&finalPriceView=1&isShowMobileAppPrice=false&cid=-999&numberOfBedrooms=&familyMode=false&adults=1&children=0&rooms=1&maxRooms=0&checkIn=2027-04-30&isCalendarCallout=false&childAges=&numberOfGuest=0&missingChildAges=false&travellerType=0&showReviewSubmissionEntry=false&currencyCode=MYR&isFreeOccSearch=false&flightSearchCriteria=%5Bobject+Object%5D&tspTypes=16&los=2&searchrequestid=cf0052c1-5c93-4727-891c-ba1f80f2d584&ds=wkEf3JY%2Fgn%2B40%2Baz | Booking.com: booking.com/hotel/my/mercure-miri-city-centre.html?label=gog235jc-10CAEoggI46AdIM1gDaKEBiAEBmAEzuAEXyAEM2AED6AEB-AEBiAIBqAIBuALP8_rRBsACAdICJDhlN2IyOTNiLTg5YTQtNDZiZi04YTc0LTNkZjE1MWMyOGMxZdgCAeACAQ&aid=304142&ucfs=1&arphpl=1&checkin=2027-04-30&checkout=2027-05-02&dest_id=9464560&dest_type=hotel&group_adults=1&req_adults=1&no_rooms=1&group_children=0&req_children=0&hpos=1&hapos=1&sr_order=popularity&srpvid=3ece7ea820b108ed&srepoch=1782496851&all_sr_blocks=946456005_367247397_2_42_0&highlighted_blocks=946456005_367247397_2_42_0&matching_block_id=946456005_367247397_2_42_0&sr_pri_blocks=946456005_367247397_2_42_0__47320&from_sustainable_property_sr=1&from=searchresults
    
    GRAB CAR FARE FROM MIRI AIRPORT (MYY) DIRECT TO VENUE:
    - Distance: ~9.5 km (~15 mins drive).
    - Estimated GrabCar Price: Between RM 11.00 to RM 18.00 for a standard 4-seater.
    
    WEDDING GIFTS & CONTRIBUTIONS POLICY:
    - If a guest asks about wedding gifts, presents, cash gifts, or contribution options: State warmly that your presence is the greatest gift, gifts are absolutely not necessary, and we are just excited to celebrate together (the more the merrier!).
    - CRITICAL RULE: If they explicitly ask about gifts, contributions, or how to give, you MUST append the text "[SHOW_QR_CODE]" at the very end of your response so the system can show them the QR option safely.
    
    CRITICAL ANSWERING RULES:
    - Always provide the full explicit links (e.g., https://www.agoda.com) when a user asks for bookings or locations.
    - Organize hotel options as clean, easy-to-read list profiles.
    - Keep answers warm, friendly, and brief (maximum 4 sentences if listing options).
    - If a guest asks a question that is not covered in your wedding information data deck, politely state that you don't have that detail and direct them to choose either "Contact Akmal" or "Contact Musfirah" using the quick action contact buttons directly below the chat screen.
    - If a guest asks how to RSVP, register, change attendance, or confirm their seat, do not explain the steps. Instead, respond with exactly: "I've opened the RSVP form on your screen for you! Please fill it out directly to secure your spot. [TRIGGER_EXISTING_RSVP]"
    - If a guest asks to CHANGE, EDIT, UPDATE, MODIFY, or CANCEL their RSVP details (such as changing their guest headcount, updating their email, or altering seating arrangements after submitting), politely state that submitted RSVPs cannot be changed automatically by the AI. Instruct them to reach out to the family directly using the coordinator buttons below so their details can be updated manually. You MUST include the words "contact" or "coordinator" in your response so the system opens the WhatsApp panel for them.
  `;

  const payload = {
    contents: [{
      parts: [{
        text: `${systemInstruction}\n\nGuest's Question: "${userPrompt}"`
      }]
    }],
    safetySettings: [
      { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
    ]
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    // 🌟 DETECT QUOTA OR RATE LIMITS (HTTP 429)
    if (responseCode === 429) {
      return "Hold up 😭 I’ve reached my limit. My owner needs to feed me more API credits! 👥 Please try clicking your option again in a few seconds, or tap one of our human coordinators below for instant help.";
    }

    const json = JSON.parse(responseText);
    
    // Handle error paylods wrapped inside an OK response status
    if (json.error && (json.error.code === 429 || json.error.status === "RESOURCE_EXHAUSTED")) {
      return "Hold up 😭 I’ve reached my limit. My owner needs to feed me more API credits! 👥 Please try clicking your option again in a few seconds, or tap one of our human coordinators below for instant help.";
    }

    if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0].text) {
      return json.candidates[0].content.parts[0].text.trim();
    } else {
      return "I couldn't process that request right now. Please try asking about hotel choices, location maps, or flight options!";
    }
  } catch (err) {
    // Catch-all for complete connection dropouts
    if (err.toString().includes("429")) {
      return "Hold up 😭 I’ve reached my limit. My owner needs to feed me more API credits! 👥 Please try clicking your option again in a few seconds, or tap one of our human coordinators below for instant help.";
    }
    return "Could not connect to the travel desk: " + err.toString();
  }
}