# Captain Emad Customs Bot - Requirements

## Bot Info
- Bot Name: Captain Emad Customs Bot
- Admin: Captain Emad
- Admin Telegram ID: 1280808746

## Phase 1 - DONE ✅
- /start sends welcome message in English
- Welcome Message:
Welcome to Captain Emad Customs Bot.
After the trial and subscription for 150 EGP, you will be able to search for the customs item, inquire about GS1, and calculate customs.
Send /start to continue.
Your Code: USR-XXXX
Waiting for acceptance from Captain Emad.
- Each user gets unique code: USR-XXXX
- User status: pending / accepted / rejected
- Admin receives new user request with buttons: Accept / Reject
- Rejected user is blocked for 5 minutes, then can send /start again
- /users : list all users with code, name, username, id, status
- /stats : total users, accepted, pending, rejected
- /delete USR-XXXX : delete a user
- Accepted user sends inquiry -> Bot replies: Received. Captain Emad will reply to you.
- Inquiry is forwarded to Admin with user code and details

## Phase 2 - TODO ⏳
- Automatic Customs Item Search: user sends HS Code or item name, bot returns customs item details
- GS1 Inquiry: user sends barcode / product, bot returns GS1 info
- Customs Calculator: calculate approximate customs based on item value, quantity, HS code
- Must work only for accepted users
- Admin bypass: Admin should not need approval, should not be counted as user

## Rules
- Only accepted users can use bot features
- Pending users only see waiting message
- Rejected users see rejection message and can retry after 5 minutes
- All user data stored in users.json
