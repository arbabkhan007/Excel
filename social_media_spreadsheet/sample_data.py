"""
Demo data.

A small handmade-ceramics brand with five social accounts, ten weeks of
content, captions, hashtags, campaigns, influencers, analytics and post
level performance.  Everything is written into the input (cream) cells so
the dashboards, reports and alerts show meaningful numbers the moment the
workbook is opened, and everything is dated relative to the build date.

Switch it all off with ``--blank``.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def month_key(day: _dt.date) -> str:
    return day.strftime("%Y-%m")


@dataclass
class Sample:
    accounts: List[Dict[str, Any]] = field(default_factory=list)
    personas: List[Dict[str, Any]] = field(default_factory=list)
    pillars: List[Dict[str, Any]] = field(default_factory=list)
    calendar: List[Dict[str, Any]] = field(default_factory=list)
    ideas: List[Dict[str, Any]] = field(default_factory=list)
    captions: List[Dict[str, Any]] = field(default_factory=list)
    hashtags: List[Dict[str, Any]] = field(default_factory=list)
    instagram: List[Dict[str, Any]] = field(default_factory=list)
    tiktok: List[Dict[str, Any]] = field(default_factory=list)
    pinterest: List[Dict[str, Any]] = field(default_factory=list)
    youtube: List[Dict[str, Any]] = field(default_factory=list)
    facebook: List[Dict[str, Any]] = field(default_factory=list)
    campaigns: List[Dict[str, Any]] = field(default_factory=list)
    launches: List[Dict[str, Any]] = field(default_factory=list)
    seasonal: List[Dict[str, Any]] = field(default_factory=list)
    influencers: List[Dict[str, Any]] = field(default_factory=list)
    ugc: List[Dict[str, Any]] = field(default_factory=list)
    repurposing: List[Dict[str, Any]] = field(default_factory=list)
    competitors: List[Dict[str, Any]] = field(default_factory=list)
    trends: List[Dict[str, Any]] = field(default_factory=list)
    analytics: List[Dict[str, Any]] = field(default_factory=list)
    performance: List[Dict[str, Any]] = field(default_factory=list)
    tests: List[Dict[str, Any]] = field(default_factory=list)
    ads: List[Dict[str, Any]] = field(default_factory=list)
    assets: List[Dict[str, Any]] = field(default_factory=list)
    goals: List[Dict[str, Any]] = field(default_factory=list)
    budget: List[Dict[str, Any]] = field(default_factory=list)
    team: List[Dict[str, Any]] = field(default_factory=list)
    tasks: List[Dict[str, Any]] = field(default_factory=list)


_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (once) the demo dataset for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg)
    return _CACHE[key]


# ---------------------------------------------------------------------------
# The content plan.  (days from today, platform, format, pillar, title)
# Negative days are already published, positive days are still to come.
# ---------------------------------------------------------------------------

PLAN = [
    (-56, "Instagram", "Carousel", "Educate",
     "5 ways to style a single stem vase"),
    (-55, "Pinterest", "Standard Pin", "Educate",
     "Autumn table styling checklist"),
    (-54, "TikTok", "Reel", "Behind the scenes",
     "Throwing a mug from start to finish"),
    (-52, "Instagram", "Reel", "Inspire",
     "Morning light in the studio"),
    (-50, "Facebook", "Feed post", "Community",
     "Meet the makers behind your mugs"),
    (-49, "Instagram", "Story", "Promote",
     "Last 12 speckled bowls of the batch"),
    (-48, "YouTube", "Long video", "Educate",
     "How a glaze is mixed - full studio tour"),
    (-47, "Pinterest", "Standard Pin", "Inspire",
     "Neutral shelf styling in 6 steps"),
    (-45, "TikTok", "Short video", "Behind the scenes",
     "Kiln opening - did they survive?"),
    (-44, "Instagram", "Feed post", "Community",
     "Your kitchen shelfies - UGC round up"),
    (-42, "Instagram", "Carousel", "Educate",
     "Stoneware vs porcelain - which to buy"),
    (-41, "Facebook", "Photo", "Promote",
     "Autumn collection now in the shop"),
    (-40, "Pinterest", "Video Pin", "Educate",
     "Three ways to fold a linen napkin"),
    (-38, "TikTok", "Reel", "Entertain",
     "Pottery fails you asked to see"),
    (-37, "Instagram", "Reel", "Promote",
     "Unboxing the autumn table set"),
    (-35, "YouTube", "Short video", "Educate",
     "Why handmade mugs keep tea hotter"),
    (-34, "Instagram", "Story", "Community",
     "This or that - glaze colour vote"),
    (-33, "Pinterest", "Standard Pin", "Promote",
     "The autumn table set - shop the look"),
    (-31, "Instagram", "Carousel", "Inspire",
     "A slow Sunday table, styled three ways"),
    (-30, "Facebook", "Feed post", "Educate",
     "Caring for handmade ceramics"),
    (-28, "TikTok", "Reel", "Behind the scenes",
     "Packing orders at 6am"),
    (-27, "Instagram", "Feed post", "Personal",
     "Why I left my desk job for clay"),
    (-25, "Pinterest", "Standard Pin", "Educate",
     "Small kitchen storage ideas"),
    (-24, "Instagram", "Reel", "Educate",
     "The 3 mugs every home needs"),
    (-23, "YouTube", "Long video", "Behind the scenes",
     "A week in a one woman ceramics studio"),
    (-21, "Instagram", "Carousel", "Community",
     "Answering your 9 most asked questions"),
    (-20, "TikTok", "Short video", "Inspire",
     "Satisfying trimming sounds"),
    (-19, "Facebook", "Link post", "Promote",
     "Restock alert - speckled dinnerware"),
    (-17, "Instagram", "Reel", "Promote",
     "Restock in 60 seconds"),
    (-16, "Pinterest", "Video Pin", "Inspire",
     "Styling a console table for autumn"),
    (-14, "Instagram", "Story", "Promote",
     "Studio sale - 20% off seconds"),
    (-13, "TikTok", "Reel", "Educate",
     "What 'seconds' means in pottery"),
    (-12, "Instagram", "Carousel", "Educate",
     "How to photograph your table setting"),
    (-10, "YouTube", "Short video", "Entertain",
     "Clay or chocolate? guess the sound"),
    (-9, "Instagram", "Feed post", "Community",
     "Customer story - Hannah's wedding set"),
    (-7, "Instagram", "Reel", "Inspire",
     "Autumn table reveal"),
    (-6, "Pinterest", "Standard Pin", "Educate",
     "Gift guide - presents for home lovers"),
    (-5, "TikTok", "Reel", "Promote",
     "Gift sets are live"),
    (-4, "Facebook", "Feed post", "Community",
     "Local market this weekend"),
    (-3, "Instagram", "Carousel", "Promote",
     "The gift guide - 8 ideas under 60"),
    (-2, "Instagram", "Story", "Community",
     "Ask me anything about ordering"),
    (-1, "TikTok", "Short video", "Behind the scenes",
     "Wrapping a gift order"),
    (0, "Instagram", "Reel", "Educate",
     "Three table settings, one dinner set"),
    (1, "Pinterest", "Standard Pin", "Promote",
     "Holiday gift guide pin"),
    (2, "Instagram", "Carousel", "Inspire",
     "Colour palette: oat, clay and sage"),
    (3, "YouTube", "Long video", "Educate",
     "Beginner's guide to buying handmade"),
    (4, "TikTok", "Reel", "Entertain",
     "Guess the glaze - part two"),
    (5, "Instagram", "Feed post", "Promote",
     "Holiday shipping dates"),
    (6, "Facebook", "Feed post", "Promote",
     "Holiday gift guide is live"),
    (8, "Instagram", "Reel", "Behind the scenes",
     "Glazing the holiday collection"),
    (9, "Pinterest", "Video Pin", "Educate",
     "How to set a holiday table"),
    (11, "Instagram", "Carousel", "Community",
     "Your holiday tables - call for photos"),
    (13, "TikTok", "Reel", "Promote",
     "Last posting day countdown"),
    (15, "Instagram", "Story", "Promote",
     "Final orders before the break"),
]

#: campaign for a plan row, by index range
CAMPAIGN_FOR = [
    (0, 17, "Autumn Table Styling"),
    (18, 29, "Studio Restock"),
    (30, 54, "Holiday Gift Guide"),
]

OBJECTIVES = {
    "Educate": "Brand authority",
    "Inspire": "Awareness",
    "Behind the scenes": "Community",
    "Community": "Engagement",
    "Promote": "Sales",
    "Entertain": "Awareness",
    "Personal": "Community",
}

CTAS = {
    "Educate": "Save this",
    "Inspire": "Follow for more",
    "Behind the scenes": "Comment below",
    "Community": "Tag a friend",
    "Promote": "Shop now",
    "Entertain": "Share with a friend",
    "Personal": "Reply to this",
}

FUNNEL = {
    "Educate": "Consideration",
    "Inspire": "Awareness",
    "Behind the scenes": "Awareness",
    "Community": "Retention",
    "Promote": "Conversion",
    "Entertain": "Awareness",
    "Personal": "Awareness",
}

SLOTS = {
    "Instagram": "Morning",
    "TikTok": "Evening",
    "Pinterest": "Lunchtime",
    "YouTube": "Afternoon",
    "Facebook": "Morning",
}

OWNERS = ["Jordan Avery", "Priya Nair", "Sam Okonkwo", "Elena Dias"]


# ---------------------------------------------------------------------------

def _build(cfg: Config) -> Sample:
    rnd = random.Random(20261011)
    today = cfg.today
    s = Sample()
    brand = cfg.brand.name

    # ------------------------------------------------------------- accounts
    account_plan = [
        ("Instagram", "Wildflower Studio", "@wildflower.studio",
         "instagram.com/wildflower.studio", "Business account", 8420,
         "Main shop window - product, styling and community", 5),
        ("TikTok", "Wildflower Studio", "@wildflowerstudio",
         "tiktok.com/@wildflowerstudio", "Creator account", 5260,
         "Reach new people with process and personality", 4),
        ("Pinterest", "Wildflower Studio", "wildflowerstudio",
         "pinterest.com/wildflowerstudio", "Business account", 3180,
         "Evergreen search traffic to the shop", 3),
        ("YouTube", "Wildflower Studio", "@wildflowerstudio",
         "youtube.com/@wildflowerstudio", "Channel", 1240,
         "Long form studio films and buying guides", 1),
        ("Facebook", "Wildflower Studio", "wildflowerstudiohome",
         "facebook.com/wildflowerstudiohome", "Page", 2640,
         "Local customers, markets and events", 2),
    ]
    for i, (platform, name, handle, url, kind, followers, purpose,
            per_week) in enumerate(account_plan):
        s.accounts.append(dict(
            platform=platform, account=name, handle=handle, url=url,
            type=kind, owner=OWNERS[i % len(OWNERS)], purpose=purpose,
            audience="Home lovers 28-45" if i != 3 else "Makers and buyers",
            followers=followers, launched=d(today, -720 + i * 90),
            target_per_week=per_week,
            bio_ok="Yes", photo_ok="Yes", cover_ok="Yes" if i != 1 else "No",
            link_ok="Yes", contact_ok="Yes" if i != 3 else "No",
            branding_ok="Yes", access="Owner + manager, 2FA on",
            notes="Link in bio points at the shop" if i == 0 else ""))

    # ------------------------------------------------------------- personas
    persona_plan = [
        ("First-home Freya", "28-34", "City flats, UK & US",
         "Slow living, thrifting, cooking for friends",
         "Her flat feels unfinished and everything matches too much",
         "A home that feels personal without spending a fortune",
         "Wants pieces with a story", "Worries handmade means fragile",
         "Instagram", "Carousel", "Awareness"),
        ("Gifting Grace", "35-45", "Suburban US & Canada",
         "Hosting, gift giving, seasonal decorating",
         "Finding a present that is not generic, close to the deadline",
         "To be the person who gives the thoughtful gift",
         "Buys for others, values packaging",
         "Will it arrive in time?", "Pinterest", "Standard Pin",
         "Conversion"),
        ("Interiors Isla", "30-42", "Design-led cities",
         "Interiors, ceramics, neutral palettes, slow design",
         "Mass produced homeware feels soulless",
         "A considered home she is proud to photograph",
         "Follows makers, buys a piece at a time",
         "Is the colour true to the photo?", "Instagram", "Reel",
         "Consideration"),
    ]
    for (name, age, place, interests, pain, goal, motivation, objection,
         platform, fmt, stage) in persona_plan:
        s.personas.append(dict(
            name=name, age=age, location=place, interests=interests,
            pain=pain, goal=goal, motivation=motivation,
            objection=objection, platform=platform, format=fmt,
            stage=stage,
            engagement="Replies to stories, saves carousels",
            faq="How long does delivery take?",
            language="'shelfie', 'tablescape', 'slow made'",
            awareness_idea="Styling inspiration reels",
            consideration_idea="Material and care explainers",
            conversion_idea="Shop the look carousels",
            retention_idea="Customer features and care tips",
            notes="Highest spend per order" if name == "Gifting Grace"
                  else ""))

    # ------------------------------------------------------------- pillars
    pillar_plan = [
        ("Educate", 0.25, "Teach one useful thing about living with "
                          "handmade pieces", "Brand authority",
         "Carousel", "How to, explainers, care guides"),
        ("Inspire", 0.20, "Show the feeling of a slow, styled home",
         "Awareness", "Reel", "Styling, light, seasonal tables"),
        ("Behind the scenes", 0.18, "Let people watch the making",
         "Community", "Reel", "Studio days, kiln openings, packing"),
        ("Community", 0.12, "Talk with people, not at them",
         "Engagement", "Story", "Questions, votes, customer features"),
        ("Promote", 0.15, "Ask for the sale, clearly and kindly",
         "Sales", "Carousel", "Launches, restocks, gift guides"),
        ("Entertain", 0.05, "Make someone smile and share it",
         "Awareness", "Reel", "Studio bloopers, trends, satisfying clips"),
        ("Personal", 0.05, "The person behind the brand", "Community",
         "Feed post", "Story of the maker, lessons, opinions"),
    ]
    for name, share, purpose, objective, fmt, examples in pillar_plan:
        s.pillars.append(dict(
            name=name, target=share, purpose=purpose, objective=objective,
            format=fmt, examples=examples,
            persona="First-home Freya" if name == "Educate"
                    else "Interiors Isla",
            notes=""))

    # ------------------------------------------------------------- captions
    caption_plan = [
        ("Single stem styling", "Educate",
         "One stem. Three minutes. A table that looks considered.",
         "Save this for your next dinner", "Save this"),
        ("Kiln opening", "Behind the scenes",
         "Opening the kiln never stops being terrifying.",
         "Tell me which glaze wins", "Comment below"),
        ("Autumn collection live", "Promote",
         "The autumn table is in the shop, and there are 40 pieces.",
         "Shop the collection", "Shop now"),
        ("Stoneware vs porcelain", "Educate",
         "Stoneware chips less. Porcelain is lighter. Here is how to "
         "choose.", "Save for your next order", "Save this"),
        ("Customer shelfie", "Community",
         "Your shelves are better styled than mine, honestly.",
         "Tag us in yours", "Tag a friend"),
        ("Why I left my desk job", "Personal",
         "Three years ago I was answering emails. Now I answer to a kiln.",
         "Tell me your leap story", "Reply to this"),
        ("Restock alert", "Promote",
         "The speckled dinnerware is back, and it goes quickly.",
         "Link in bio", "Link in bio"),
        ("Care guide", "Educate",
         "Handmade does not mean delicate. It means made by hands.",
         "Save the care guide", "Save this"),
        ("Gift guide", "Promote",
         "Eight presents for the person whose home always smells of "
         "coffee.", "Shop the guide", "Shop now"),
        ("Slow Sunday", "Inspire",
         "Nothing on the calendar, everything on the table.",
         "Follow for slow living", "Follow for more"),
        ("Studio week", "Behind the scenes",
         "Six days, 180 mugs, one very tired kiln.",
         "Watch the full film", "Watch to the end"),
        ("Holiday shipping", "Promote",
         "Last posting dates are confirmed - here they are.",
         "Order before Friday", "Shop now"),
    ]
    for i, (title, pillar, body, cta_line, cta) in enumerate(caption_plan):
        platform = ["Instagram", "TikTok", "Pinterest", "Facebook"][i % 4]
        hook = body.split(".")[0] + "."
        full = f"{hook}\n\n{body}\n\n{cta_line} {cfg.brand.hashtag}"
        s.captions.append(dict(
            title=title, platform=platform, pillar=pillar, hook=hook,
            body=body, cta=cta_line, full=full,
            short=hook, long=full + "\n\nMade slowly in the studio.",
            persona="First-home Freya" if i % 2 else "Interiors Isla",
            tone="Warm" if i % 3 else "Playful",
            emojis="One or two, never a wall",
            group=f"Group {1 + i % 4}", cta_type=cta,
            offer="20% off seconds" if "sale" in title.lower() else "",
            approval="Approved" if i < 9 else "Awaiting review",
            created=d(today, -70 + i * 5), edited=d(today, -14 + i),
            reusable="Yes" if i % 3 else "No",
            notes="Best performing hook so far" if i == 0 else ""))

    # ------------------------------------------------------------ hashtags
    tag_plan = [
        ("#handmadeceramics", "Broad reach", "Educate", 860000),
        ("#slowliving", "Broad reach", "Inspire", 1200000),
        ("#tablescape", "Niche", "Inspire", 240000),
        ("#potterylife", "Community", "Behind the scenes", 410000),
        ("#wildflowerstudio", "Branded", "Community", 1800),
        ("#handmademug", "Niche", "Promote", 96000),
        ("#stoneware", "Niche", "Educate", 150000),
        ("#kitchenshelfie", "Community", "Community", 72000),
        ("#giftsforhomelovers", "Seasonal", "Promote", 18000),
        ("#autumntable", "Seasonal", "Inspire", 34000),
        ("#madeinthestudio", "Branded", "Behind the scenes", 2400),
        ("#ceramicartist", "Broad reach", "Behind the scenes", 540000),
        ("#smallbatchpottery", "Niche", "Promote", 21000),
        ("#neutralhome", "Broad reach", "Inspire", 680000),
        ("#shopsmall", "Community", "Promote", 2100000),
        ("#glazetest", "Niche", "Behind the scenes", 42000),
        ("#dinnerpartyideas", "Niche", "Educate", 88000),
        ("#holidaygiftguide", "Campaign", "Promote", 120000),
        ("#handmadehome", "Broad reach", "Inspire", 390000),
        ("#ceramicsofinstagram", "Community", "Community", 310000),
        ("#localmaker", "Location", "Community", 64000),
        ("#speckledpottery", "Product", "Promote", 9800),
        ("#mugshot", "Trending", "Entertain", 55000),
        ("#studiotour", "Niche", "Behind the scenes", 47000),
    ]
    keywords = {
        "Educate": ("how to style a table", "Informational"),
        "Inspire": ("autumn tablescape ideas", "Inspirational"),
        "Promote": ("handmade mug gift", "Transactional"),
        "Community": ("handmade home community", "Navigational"),
        "Behind the scenes": ("pottery studio tour", "Informational"),
        "Entertain": ("satisfying pottery video", "Informational"),
    }
    for i, (tag, kind, pillar, volume) in enumerate(tag_plan):
        keyword, intent = keywords.get(pillar, ("handmade ceramics",
                                                "Informational"))
        s.hashtags.append(dict(
            tag=tag, type=kind, pillar=pillar,
            topic=pillar.lower(),
            platform=["Instagram", "TikTok", "Pinterest", "Facebook"][i % 4],
            group=f"Group {1 + i % 4}",
            keyword=keyword, related=f"{keyword}, {pillar.lower()} ideas",
            intent=intent,
            difficulty=["Low", "Medium", "High"][i % 3],
            volume=volume, uses=rnd.randint(2, 28),
            last_used=d(today, -rnd.randint(1, 40)),
            banned="No",
            notes="Branded - always include" if kind == "Branded" else ""))

    # ------------------------------------------------------------- calendar
    def campaign_for(index: int) -> str:
        for start, end, name in CAMPAIGN_FOR:
            if start <= index <= end:
                return name
        return ""

    caption_titles = [c["title"] for c in s.captions]
    asset_ids = [f"AS-{i + 1:04d}" for i in range(14)]

    for i, (offset, platform, fmt, pillar, title) in enumerate(PLAN):
        if offset < 0:
            status = "Published"
        elif offset <= 2:
            status = "Scheduled"
        elif offset <= 6:
            status = "Approved"
        elif offset <= 9:
            status = "Draft"
        elif offset <= 12:
            status = "In production"
        else:
            status = "Brief"
        day = d(today, offset)
        s.calendar.append(dict(
            date=day, slot=SLOTS.get(platform, "Morning"),
            platform=platform, title=title,
            description=f"{pillar} post for {platform.lower()}: {title}.",
            pillar=pillar, format=fmt, campaign=campaign_for(i),
            objective=OBJECTIVES.get(pillar, "Awareness"),
            funnel=FUNNEL.get(pillar, "Awareness"),
            persona=["First-home Freya", "Gifting Grace",
                     "Interiors Isla"][i % 3],
            caption=caption_titles[i % len(caption_titles)],
            hashtag_group=f"Group {1 + i % 4}",
            asset=asset_ids[i % len(asset_ids)],
            cta=CTAS.get(pillar, "Follow for more"),
            owner=OWNERS[i % len(OWNERS)],
            status=status,
            url=(f"{platform.lower()}.com/{brand.lower().replace(' ', '')}"
                 f"/p/{abs(offset):02d}{i:02d}") if offset < 0 else "",
            published="Yes" if offset < 0 else "No",
            notes="Top performer" if i in (3, 24, 35) else ""))

    # --------------------------------------------------------- performance
    # Only published posts are measured.  Engagement is loosely driven by the
    # platform so the comparison tables have something to show.
    base_reach = {"Instagram": 4200, "TikTok": 9800, "Pinterest": 2600,
                  "YouTube": 1500, "Facebook": 1400}
    published = [(i, row) for i, row in enumerate(s.calendar)
                 if row["status"] == "Published"]
    for i, row in published:
        platform = row["platform"]
        spread = rnd.uniform(0.55, 1.85)
        reach = int(base_reach[platform] * spread)
        impressions = int(reach * rnd.uniform(1.15, 1.6))
        rate = rnd.uniform(0.028, 0.085)
        if row["pillar"] in ("Educate", "Community"):
            rate *= 1.25
        if row["pillar"] == "Promote":
            rate *= 0.75
        engagements = max(6, int(reach * rate))
        likes = int(engagements * 0.72)
        comments = max(1, int(engagements * 0.09))
        shares = max(1, int(engagements * 0.07))
        saves = engagements - likes - comments - shares
        clicks = int(reach * rnd.uniform(0.004, 0.022))
        conversions = int(clicks * rnd.uniform(0.02, 0.09))
        s.performance.append(dict(
            content=f"C-{i + 1:04d}", reach=reach, impressions=impressions,
            likes=likes, comments=comments, shares=shares,
            saves=max(0, saves), clicks=clicks, conversions=conversions,
            revenue=round(conversions * rnd.uniform(34, 72), 2),
            watch=(round(rnd.uniform(4.5, 31.0), 1)
                   if platform in ("TikTok", "YouTube", "Instagram")
                   else 0),
            target_rate=0.045,
            notes="Hook tested in A/B 2" if i == 35 else ""))

    # ------------------------------------------------- platform planners --
    def content_of(platform: str, limit: int) -> List[tuple]:
        out = []
        for i, row in enumerate(s.calendar):
            if row["platform"] == platform:
                out.append((f"C-{i + 1:04d}", row))
            if len(out) >= limit:
                break
        return out

    for cid, row in content_of("Instagram", 9):
        reel = row["format"] in ("Reel", "Story")
        s.instagram.append(dict(
            content=cid, format=row["format"], date=row["date"],
            hook="Open on the finished table, then rewind" if reel
                 else "Strong first slide question",
            slides=0 if reel else rnd.choice([5, 6, 7, 8, 10]),
            duration=rnd.choice([12, 18, 24, 30]) if reel else 0,
            cover="covers/ig-" + cid.lower() + ".jpg",
            collab="No", story_sequence="Hook > proof > poll > link"
            if row["format"] == "Story" else "",
            profile_visits=rnd.randint(40, 320),
            followers_gained=rnd.randint(4, 85),
            notes=""))

    for cid, row in content_of("TikTok", 7):
        s.tiktok.append(dict(
            content=cid, date=row["date"],
            hook="First line on screen in under 2 seconds",
            sound="Trending - 'slow morning' lo-fi",
            script="Hook > process > payoff > CTA",
            shots="Wheel close up, hands, finished piece",
            duration=rnd.choice([15, 21, 28, 34, 45]),
            text_on_screen="3 mugs every home needs",
            cover="covers/tt-" + cid.lower() + ".jpg",
            series="Studio diaries" if rnd.random() > 0.5 else "",
            watch_time=round(rnd.uniform(5.5, 22.0), 1),
            completion=round(rnd.uniform(0.28, 0.74), 3),
            profile_views=rnd.randint(60, 540),
            followers_gained=rnd.randint(8, 160),
            notes=""))

    for cid, row in content_of("Pinterest", 7):
        s.pinterest.append(dict(
            content=cid, date=row["date"],
            board="Autumn table styling",
            board_url="pinterest.com/wildflowerstudio/autumn-table",
            pin_title=row["title"],
            pin_description=f"{row['title']} - ideas from a small "
                            f"ceramics studio.",
            destination="shop.wildflowerstudio.com/autumn",
            image="pins/" + cid.lower() + ".png",
            dimensions="1000 x 1500",
            keyword="autumn tablescape ideas",
            intent="Inspirational",
            seasonal="Yes" if row["pillar"] == "Promote" else "No",
            outbound_clicks=rnd.randint(12, 180),
            audience_growth=rnd.randint(5, 60),
            refresh_due=d(today, rnd.randint(30, 150)),
            notes=""))

    for cid, row in content_of("YouTube", 5):
        long_form = row["format"] == "Long video"
        s.youtube.append(dict(
            content=cid, date=row["date"],
            format="Long video" if long_form else "Short",
            title_option=row["title"],
            description="Full studio film with chapters and links.",
            script="Cold open > intro > 4 chapters > CTA",
            thumbnail="thumbs/" + cid.lower() + ".png",
            thumbnail_idea="Hands + kiln glow + 3 word title",
            hook="The kiln is at 1240 degrees and I am nervous",
            chapters="00:00 intro, 02:10 glaze, 06:40 firing",
            keywords="pottery studio tour, handmade ceramics",
            tags="pottery, ceramics, studio vlog",
            playlist="Studio films",
            end_screen="Subscribe + next film",
            watch_time=round(rnd.uniform(120, 940), 1),
            avg_duration=round(rnd.uniform(2.4, 7.8), 1),
            retention=round(rnd.uniform(0.32, 0.61), 3),
            ctr=round(rnd.uniform(0.035, 0.092), 3),
            subs_gained=rnd.randint(4, 70), subs_lost=rnd.randint(0, 6),
            traffic="Browse + suggested",
            notes=""))

    for cid, row in content_of("Facebook", 6):
        s.facebook.append(dict(
            content=cid, date=row["date"], format=row["format"],
            audience="Local + existing customers",
            targeting="Home decor interest, 40km radius",
            link="shop.wildflowerstudio.com",
            reactions=rnd.randint(18, 160), link_clicks=rnd.randint(5, 90),
            followers_gained=rnd.randint(1, 30),
            group_post="Yes" if rnd.random() > 0.7 else "No",
            event="Autumn market" if row["pillar"] == "Community" else "",
            notes=""))

    # ------------------------------------------------------------ campaigns
    campaign_plan = [
        ("Autumn Table Styling", "Awareness", -60, -32, 320.0, 286.40,
         "Autumn collection", "Running"),
        ("Studio Restock", "Sales", -31, -12, 180.0, 164.20,
         "Speckled dinnerware", "Complete"),
        ("Holiday Gift Guide", "Sales", -8, 30, 650.0, 212.75,
         "Gift sets", "Running"),
        ("Evergreen Education", "Brand authority", -120, 180, 0.0, 0.0,
         "All products", "Running"),
    ]
    for (name, objective, start, end, budget, spend, product,
         status) in campaign_plan:
        s.campaigns.append(dict(
            name=name, objective=objective, description=(
                f"{objective} campaign for {product.lower()}."),
            start=d(today, start), end=d(today, end),
            audience="First-home Freya" if objective != "Sales"
                     else "Gifting Grace",
            message="Slow made pieces for a table you actually use",
            product=product,
            platforms="Instagram, TikTok, Pinterest",
            owner=OWNERS[0], budget=budget, spend=spend,
            offer="Free shipping over 75" if objective == "Sales" else "",
            landing="shop.wildflowerstudio.com/" +
                    name.lower().replace(" ", "-"),
            cta="Shop now" if objective == "Sales" else "Follow for more",
            utm="utm_source=social&utm_campaign=" +
                name.lower().replace(" ", "_"),
            leads=rnd.randint(18, 90), conversions=rnd.randint(6, 48),
            revenue=round(rnd.uniform(400, 2600), 2),
            status=status, notes=""))

    # ------------------------------------------------------------- launches
    launch_plan = [
        ("Autumn Table Set", -40, "Complete"),
        ("Holiday Gift Boxes", 12, "In production"),
    ]
    for product, offset, status in launch_plan:
        launch = d(today, offset)
        s.launches.append(dict(
            product=product, launch=launch,
            campaign="Autumn Table Styling" if offset < 0
                     else "Holiday Gift Guide",
            pre_start=d(today, offset - 14), post_end=d(today, offset + 14),
            teaser="Three teaser stories + one reel",
            announcement="Carousel + email on launch morning",
            demo="Reel showing the set on a real table",
            benefits="Dishwasher safe, stacks, made in small batches",
            testimonial="Hannah's wedding set story",
            offer="Early access for the email list",
            code="EARLY10", code_uses=rnd.randint(8, 60),
            reminder=d(today, offset - 1),
            email="Teaser, launch, last call",
            influencer="Mara Lindqvist",
            budget=420.0, engagement=rnd.randint(400, 2400),
            leads=rnd.randint(20, 110), sales=rnd.randint(12, 90),
            revenue=round(rnd.uniform(900, 4200), 2),
            status=status,
            review="Sold out in 6 days - make 30% more next time"
                   if offset < 0 else ""))

    # ------------------------------------------------------------- seasonal
    year = cfg.year
    seasonal_plan = [
        (_dt.date(year, 1, 1), "New Year", "Public holiday", "Fresh start"),
        (_dt.date(year, 2, 14), "Valentine's Day", "Cultural moment",
         "Gifts for two"),
        (_dt.date(year, 3, 20), "First day of spring", "Cultural moment",
         "Spring table"),
        (_dt.date(year, 5, 10), "Mother's Day", "Cultural moment",
         "Gift guide"),
        (_dt.date(year, 6, 21), "Father's Day", "Cultural moment",
         "Mugs for dads"),
        (_dt.date(year, 8, 20), "Back to school", "Seasonal sale",
         "Lunch and storage"),
        (_dt.date(year, 9, 22), "First day of autumn", "Cultural moment",
         "Autumn table launch"),
        (_dt.date(year, 10, 31), "Halloween", "Public holiday",
         "Spooky glaze fun"),
        (_dt.date(year, 11, 27), "Black Friday", "Seasonal sale",
         "Best offer of the year"),
        (_dt.date(year, 11, 30), "Cyber Monday", "Seasonal sale",
         "Online only bundle"),
        (_dt.date(year, 12, 12), "Last posting day", "Industry event",
         "Shipping deadline reminders"),
        (_dt.date(year, 12, 25), "Christmas Day", "Public holiday",
         "Thank you post"),
    ]
    for day, name, kind, angle in seasonal_plan:
        s.seasonal.append(dict(
            date=day, name=name, type=kind, angle=angle,
            campaign="Holiday Gift Guide" if day.month in (11, 12) else "",
            pillar="Promote" if kind == "Seasonal sale" else "Inspire",
            platform="Instagram, Pinterest",
            lead_days=14 if kind == "Seasonal sale" else 7,
            hashtags="#holidaygiftguide" if day.month in (11, 12)
                     else "#slowliving",
            budget=250.0 if kind == "Seasonal sale" else 0.0,
            status="Planned", notes=""))

    # ---------------------------------------------------------- influencers
    influencer_plan = [
        ("Mara Lindqvist", "@maramakes", 24800, 0.052, "Gifted product",
         "Agreed", 0.0, -20),
        ("Tom Whitfield", "@slowhomeuk", 61200, 0.031, "Paid post",
         "Complete", 450.0, -38),
        ("Aisha Rahman", "@thehostingedit", 15400, 0.067, "Affiliate",
         "Negotiating", 0.0, -6),
        ("Nina Castro", "@ceramicdaily", 92000, 0.024, "Account takeover",
         "Contacted", 0.0, -3),
    ]
    for (name, handle, followers, rate, kind, status, fee,
         offset) in influencer_plan:
        s.influencers.append(dict(
            name=name, handle=handle,
            url=f"instagram.com/{handle.strip('@')}",
            contact=f"{handle.strip('@')}@email.com",
            niche="Interiors and slow living", followers=followers,
            engagement=rate, demographic="78% women, 25-44, UK & US",
            type=kind, campaign="Holiday Gift Guide",
            outreach=d(today, offset), status=status,
            response="Replied in 2 days" if status != "Contacted" else "",
            rate=fee, deliverables="1 reel + 3 stories + 1 grid post",
            deadline=d(today, offset + 30),
            approval="Approved" if status == "Complete"
                     else "Awaiting review",
            code=f"{name.split()[0].upper()}10",
            cost=fee + (60.0 if kind == "Gifted product" else 0.0),
            leads=rnd.randint(0, 70), sales=rnd.randint(0, 40),
            revenue=round(rnd.uniform(0, 2200), 2), notes=""))

    # ------------------------------------------------------------------ ugc
    ugc_plan = [
        ("Hannah Price", "Tagged post", "Instagram", "Wedding dinner set",
         5, "Granted"),
        ("Diego Marin", "Story mention", "Instagram", "Speckled mug",
         5, "Granted"),
        ("Keira O'Neill", "Product review", "Other", "Breakfast set",
         4, "Granted"),
        ("Yusuf Demir", "Hashtag", "TikTok", "Plant pots", 5, "Requested"),
        ("Lucy Braun", "Competition entry", "Instagram", "Serving bowl",
         4, "Granted"),
        ("Omar Haddad", "Direct message", "Instagram", "Gift box", 5,
         "Not requested"),
    ]
    for i, (person, source, platform, product, score,
            permission) in enumerate(ugc_plan):
        s.ugc.append(dict(
            person=person, source=source, platform=platform,
            url=f"{platform.lower()}.com/p/ugc{i + 1:03d}",
            type="Photo" if i % 2 else "Video clip", product=product,
            quote="It is the first mug I reach for every single morning.",
            score=score, permission=permission,
            rights="Social use, 12 months" if permission == "Granted"
                   else "",
            expires=d(today, 330) if permission == "Granted" else None,
            attribution=f"Credit @{person.split()[0].lower()}",
            approval="Approved" if permission == "Granted"
                     else "Awaiting review",
            repost=d(today, 3 + i * 2),
            destination="Instagram Story + Facebook",
            engagement=rnd.randint(120, 900), notes=""))

    # ---------------------------------------------------------- repurposing
    repurpose_sources = [2, 6, 12, 22, 24, 30, 35, 37]
    destinations = [("Pinterest", "Standard Pin"), ("Instagram", "Carousel"),
                    ("YouTube", "Short video"), ("Facebook", "Feed post"),
                    ("TikTok", "Reel"), ("Instagram", "Story"),
                    ("Pinterest", "Video Pin"), ("Instagram", "Reel")]
    for i, index in enumerate(repurpose_sources):
        row = s.calendar[index]
        platform, fmt = destinations[i % len(destinations)]
        done = i < 5
        s.repurposing.append(dict(
            source=f"C-{index + 1:04d}", destination_platform=platform,
            new_format=fmt,
            idea=f"Cut the best 20 seconds for {platform}",
            angle="Lead with the result, not the process",
            caption=f"{row['title']} - the short version",
            cta="Save this", asset=f"AS-{(i % 14) + 1:04d}",
            owner=OWNERS[(i + 1) % len(OWNERS)],
            deadline=d(today, -14 + i * 4),
            status="Published" if done else "Planned",
            new_date=d(today, -10 + i * 4) if done else d(today, 7 + i),
            new_url=(f"{platform.lower()}.com/p/re{i:03d}") if done else "",
            notes=""))

    # ----------------------------------------------------------- competitors
    competitor_plan = [
        ("Clay & Co", "Instagram", 48200, 5, "Educate + promote"),
        ("The Pottery Edit", "TikTok", 112000, 7, "Behind the scenes"),
        ("Oat Ceramics", "Pinterest", 26400, 4, "Styling inspiration"),
        ("Studio Lune", "Instagram", 18900, 3, "Personal storytelling"),
    ]
    for i, (name, platform, size, per_week, pillars) in enumerate(
            competitor_plan):
        s.competitors.append(dict(
            name=name, url=f"{platform.lower()}.com/{name.lower()}",
            platform=platform, niche="Handmade ceramics", audience=size,
            frequency=f"{per_week} posts a week", pillars=pillars,
            best_post="Kiln opening reel, 240k views",
            engagement=round(rnd.uniform(0.018, 0.061), 3),
            formats="Reels and carousels",
            campaign="Running a 12 days of gifting series"
                     if i == 0 else "",
            questions="Is it dishwasher safe?",
            trending="Glaze reveal videos",
            news="Pinterest pushing video pins harder",
            gap="Nobody explains how to choose a size",
            opportunity="Own the 'how to choose' search",
            reviewed=d(today, -7 - i * 3),
            source=f"{platform.lower()}.com/{name.lower()}",
            notes=""))

    # ---------------------------------------------------------------- trends
    trend_plan = [
        ("Slow morning lo-fi sound", "Audio", "TikTok", 9, -4),
        ("Guess the sound", "Format", "TikTok", 7, -9),
        ("Shelfie tour", "Format", "Instagram", 8, -16),
        ("12 days of gifting", "Topic", "Instagram", 9, -2),
        ("Before and after styling", "Format", "Pinterest", 6, -22),
    ]
    for name, kind, platform, score, offset in trend_plan:
        s.trends.append(dict(
            name=name, platform=platform, found=d(today, offset),
            url=f"{platform.lower()}.com/trend/{name.split()[0].lower()}",
            type=kind, audio=name if kind == "Audio" else "",
            topic=name, audience="First-home Freya", score=score,
            deadline=d(today, offset + 14),
            idea=f"Use {name.lower()} for a styling reveal",
            owner=OWNERS[1], status="Published" if offset < -10
                                    else "In production",
            posted=d(today, offset + 5) if offset < -10 else None,
            reach=rnd.randint(2000, 24000) if offset < -10 else 0,
            engagement=rnd.randint(120, 1900) if offset < -10 else 0,
            outcome="Best reach of the month" if offset < -10 else "",
            notes=""))

    # ------------------------------------------------------------- analytics
    #: three months of account level numbers per platform
    months = []
    cursor = today.replace(day=1)
    for back in (2, 1, 0):
        month = cursor
        for _ in range(back):
            month = (month - _dt.timedelta(days=1)).replace(day=1)
        months.append(month)
    growth = {"Instagram": 180, "TikTok": 260, "Pinterest": 90,
              "YouTube": 45, "Facebook": 30}
    for account in s.accounts:
        platform = account["platform"]
        followers = account["followers"] - growth[platform] * 3
        for month in months:
            gained = growth[platform] + rnd.randint(-30, 70)
            lost = rnd.randint(2, 28)
            start = followers
            followers = start + gained - lost
            posts = len([row for row in s.calendar
                         if row["platform"] == platform
                         and month_key(row["date"]) == month_key(month)
                         and row["status"] == "Published"])
            reach = int((posts or 1) * {
                "Instagram": 5200, "TikTok": 11800, "Pinterest": 3400,
                "YouTube": 1900, "Facebook": 1600}[platform]
                * rnd.uniform(0.8, 1.3))
            impressions = int(reach * rnd.uniform(1.2, 1.55))
            engagements = int(reach * rnd.uniform(0.035, 0.075))
            clicks = int(reach * rnd.uniform(0.006, 0.02))
            leads = int(clicks * rnd.uniform(0.08, 0.2))
            conversions = int(leads * rnd.uniform(0.2, 0.45))
            s.analytics.append(dict(
                month=month, platform=platform,
                account=account["account"],
                followers_start=start, followers_end=followers,
                lost=lost, posts=posts, reach=reach,
                impressions=impressions, engagements=engagements,
                likes=int(engagements * 0.74),
                comments=int(engagements * 0.1),
                shares=int(engagements * 0.08),
                saves=int(engagements * 0.08),
                profile_visits=int(reach * 0.035),
                clicks=clicks, leads=leads, conversions=conversions,
                revenue=round(conversions * rnd.uniform(38, 74), 2),
                spend=round({"Instagram": 120, "TikTok": 60,
                             "Pinterest": 40, "YouTube": 0,
                             "Facebook": 50}[platform]
                            * rnd.uniform(0.7, 1.4), 2),
                notes=""))

    # ------------------------------------------------------------ A/B tests
    test_plan = [
        ("Hook: question vs statement", "Hook", "Instagram",
         "Question hook beats a statement hook on saves",
         "Which mug do you reach for?", "This is the mug you reach for",
         -30, -23, "Version A", 0.061, 0.043),
        ("Carousel length", "Caption length", "Instagram",
         "Five slides hold attention better than ten",
         "5 slides", "10 slides", -22, -15, "Version B", 0.038, 0.052),
        ("Posting time", "Posting time", "TikTok",
         "Evening posts out-reach morning posts",
         "08:00 post", "19:00 post", -14, -7, "Version B", 0.031, 0.049),
        ("Thumbnail style", "Thumbnail", "YouTube",
         "Hands in frame lift click through rate",
         "Product only", "Hands holding piece", -6, 8, "", 0.041, 0.0),
    ]
    for (name, variable, platform, hypothesis, a, b, start, end, winner,
         a_rate, b_rate) in test_plan:
        s.tests.append(dict(
            name=name, variable=variable, platform=platform,
            hypothesis=hypothesis, version_a=a, version_b=b,
            start=d(today, start), end=d(today, end),
            metric="Engagement rate", secondary="Saves",
            audience="Same account, similar pillar",
            a_result=a_rate, b_result=b_rate,
            a_clicks=rnd.randint(20, 180), b_clicks=rnd.randint(20, 180),
            a_conversions=rnd.randint(1, 20),
            b_conversions=rnd.randint(1, 20),
            winner=winner,
            confidence="Small sample - repeat next month" if winner
                       else "Still running",
            next_test="Try the question hook on TikTok" if winner else "",
            status="Complete" if winner else "Running"))

    # ------------------------------------------------------------------ ads
    ad_plan = [
        ("Autumn collection - traffic", "Instagram", "Traffic", -58, -40,
         180.0, 174.20),
        ("Restock retargeting", "Facebook", "Retargeting", -30, -16,
         120.0, 118.65),
        ("Gift guide - sales", "Instagram", "Sales", -8, 14, 320.0, 96.40),
        ("Pin promotion", "Pinterest", "Traffic", -12, 18, 90.0, 38.20),
        ("Studio film - views", "YouTube", "Video views", -24, -10, 70.0,
         68.90),
        ("Local market - awareness", "Facebook", "Awareness", -6, 2, 45.0,
         41.10),
    ]
    for (name, platform, objective, start, end, budget, spend) in ad_plan:
        impressions = int(spend * rnd.uniform(220, 480))
        reach = int(impressions * rnd.uniform(0.55, 0.8))
        clicks = int(impressions * rnd.uniform(0.008, 0.021))
        conversions = int(clicks * rnd.uniform(0.02, 0.08))
        s.ads.append(dict(
            name=name, platform=platform, objective=objective,
            campaign="Holiday Gift Guide" if start > -20
                     else "Autumn Table Styling",
            audience="Interests: handmade, interiors, slow living",
            segment="Lookalike 2% of buyers" if "retarget" in name.lower()
                    else "Cold interest audience",
            creative=f"AS-{rnd.randint(1, 14):04d}",
            copy="Slow made pieces for a table you actually use",
            headline="Handmade ceramics, made in small batches",
            landing="shop.wildflowerstudio.com",
            start=d(today, start), end=d(today, end),
            budget=budget, spend=spend, impressions=impressions,
            reach=reach, clicks=clicks, conversions=conversions,
            revenue=round(conversions * rnd.uniform(42, 88), 2),
            notes=""))

    # ---------------------------------------------------------------- assets
    asset_plan = [
        ("Autumn table hero", "Photo", "1080 x 1350"),
        ("Studio hands close up", "Photo", "1080 x 1350"),
        ("Kiln opening clip", "Video clip", "1080 x 1920"),
        ("Glaze pour slow motion", "Video clip", "1080 x 1920"),
        ("Carousel template - teal", "Canva template", "1080 x 1350"),
        ("Carousel template - oat", "Canva template", "1080 x 1350"),
        ("Reel cover set", "Reel cover", "1080 x 1920"),
        ("Pin template - vertical", "Canva template", "1000 x 1500"),
        ("Studio film thumbnail", "Thumbnail", "1280 x 720"),
        ("Product shot - speckled mug", "Product shot", "2000 x 2000"),
        ("Product shot - dinner set", "Product shot", "2000 x 2000"),
        ("Gift box flat lay", "Photo", "1080 x 1080"),
        ("Hannah wedding set UGC", "UGC clip", "1080 x 1920"),
        ("Brand logo pack", "Logo", "Vector"),
    ]
    for i, (title, kind, size) in enumerate(asset_plan):
        s.assets.append(dict(
            title=title, type=kind,
            link=f"drive/wildflower/{title.lower().replace(' ', '-')}",
            storage="Google Drive / Brand assets",
            template="canva.com/wildflower-templates"
                     if "template" in kind.lower() else "",
            creator=OWNERS[i % len(OWNERS)],
            created=d(today, -200 + i * 11),
            updated=d(today, -40 + i * 2),
            dimensions=size,
            platform="Instagram, Pinterest" if "1350" in size
                     else "All platforms",
            approval="Approved" if i < 12 else "Awaiting review",
            rights="Owned" if kind != "UGC clip" else "UGC with permission",
            expires=d(today, 300) if kind == "UGC clip" else None,
            product="Autumn collection" if i < 5 else "Gift sets",
            campaign="Autumn Table Styling" if i < 6
                     else "Holiday Gift Guide",
            reuse=rnd.randint(1, 9),
            notes="Highest saving carousel template" if i == 4 else ""))

    # ----------------------------------------------------------------- goals
    goal_plan = [
        ("Instagram followers", "Followers", "Instagram", 7400, 9500,
         300),
        ("TikTok followers", "Followers", "TikTok", 4300, 7000, 300),
        ("Average engagement rate", "Engagement", "All platforms", 0.038,
         0.055, 180),
        ("Monthly website clicks", "Website traffic", "All platforms",
         900, 1800, 120),
        ("Email sign ups", "Email list", "All platforms", 210, 600, 240),
        ("Monthly leads", "Leads", "All platforms", 90, 240, 180),
        ("Posts published per month", "Content output", "All platforms",
         34, 52, 90),
        ("Revenue from social", "Revenue", "All platforms", 9800, 36000,
         300),
    ]
    for i, (name, category, platform, baseline, target,
            days) in enumerate(goal_plan):
        s.goals.append(dict(
            name=name, category=category, platform=platform,
            description=f"Move {name.lower()} from {baseline} to {target}.",
            baseline=baseline, target=target,
            target_date=d(today, days),
            current=round(baseline + (target - baseline)
                          * rnd.uniform(0.25, 0.78), 3),
            milestone_month=round(baseline + (target - baseline) / 12, 2),
            milestone_quarter=round(baseline + (target - baseline) / 4, 2),
            owner=OWNERS[i % len(OWNERS)],
            campaign="Holiday Gift Guide" if category == "Revenue" else "",
            plan="Content Calendar", budget=round(target * 0.02, 2),
            status="On track" if i % 3 else "At risk",
            notes=""))

    # ---------------------------------------------------------------- budget
    budget_plan = [
        ("Paid advertising", 350.0, 348.55, "Meta + Pinterest ads"),
        ("Influencer fees", 150.0, 450.0, "One paid post this quarter"),
        ("Content production", 120.0, 96.40, "Props, flowers, linens"),
        ("Design & editing", 90.0, 88.00, "Video editor, 4 hours"),
        ("Software & subscriptions", 74.0, 74.00, "Scheduler + Canva"),
        ("Contractor fees", 60.0, 0.0, "Spare capacity"),
        ("Props & samples", 40.0, 61.20, "Autumn styling props"),
        ("Training", 25.0, 0.0, "Short course fund"),
        ("Stock media", 15.0, 12.99, "Music licence"),
        ("Printing", 20.0, 18.40, "Market flyers"),
    ]
    for category, planned, actual, note in budget_plan:
        s.budget.append(dict(
            category=category, month=today.replace(day=1),
            platform="All platforms" if category != "Paid advertising"
                     else "Instagram",
            campaign="Holiday Gift Guide" if actual > 100 else "",
            planned=planned, actual=actual, notes=note))

    # ------------------------------------------------------------------ team
    team_plan = [
        ("Jordan Avery", "Owner", "jordan@wildflowerstudio.com", 12),
        ("Priya Nair", "Social media manager", "priya@wildflower.com", 20),
        ("Sam Okonkwo", "Video editor", "sam@freelance.com", 8),
        ("Elena Dias", "Copywriter", "elena@freelance.com", 6),
    ]
    for name, role, email, hours in team_plan:
        s.team.append(dict(
            name=name, role=role, contact=email,
            platform="All platforms", hours=hours,
            rate=0.0 if role == "Owner" else round(28 + hours, 2),
            reviews="Yes" if role in ("Owner", "Social media manager")
                    else "No",
            approves="Yes" if role == "Owner" else "No",
            notes=""))

    task_plan = [
        ("Write December captions", "Elena Dias", -2, "In progress",
         "High"),
        ("Edit holiday reel", "Sam Okonkwo", 1, "Not started", "High"),
        ("Shoot gift box flat lay", "Jordan Avery", 2, "Not started",
         "Medium"),
        ("Schedule week 50", "Priya Nair", 3, "Not started", "High"),
        ("Refresh top 5 pins", "Priya Nair", 6, "Not started", "Medium"),
        ("Reply to UGC permissions", "Priya Nair", -1, "In progress",
         "High"),
        ("Brief influencer collab", "Jordan Avery", 5, "Not started",
         "Medium"),
        ("Update hashtag groups", "Priya Nair", 9, "Not started", "Low"),
        ("Monthly analytics entry", "Priya Nair", 12, "Not started",
         "High"),
        ("Plan January pillars", "Jordan Avery", 16, "Not started", "Low"),
    ]
    for title, owner, offset, status, priority in task_plan:
        s.tasks.append(dict(
            task=title, owner=owner, due=d(today, offset), status=status,
            priority=priority,
            campaign="Holiday Gift Guide" if offset > 0 else "",
            content="", notes=""))

    # ------------------------------------------------------------- ideas --
    idea_plan = [
        ("The 5 mug sizes explained", "Educate", "Instagram", "Carousel", 9),
        ("A day in the studio in 60 seconds", "Behind the scenes", "TikTok",
         "Reel", 8),
        ("How to style open shelving", "Educate", "Pinterest",
         "Standard Pin", 9),
        ("Clay sourcing - where it comes from", "Educate", "YouTube",
         "Long video", 6),
        ("Customer kitchen tour", "Community", "Instagram", "Reel", 7),
        ("Glaze colour vote", "Community", "Instagram", "Story", 8),
        ("Packing an order ASMR", "Behind the scenes", "TikTok",
         "Short video", 7),
        ("What 100 hours of throwing looks like", "Personal", "Instagram",
         "Carousel", 6),
        ("Three dinner party tables", "Inspire", "Pinterest", "Video Pin",
         8),
        ("Myth: handmade is fragile", "Educate", "Instagram", "Reel", 9),
        ("Studio playlist", "Personal", "Facebook", "Feed post", 4),
        ("Seconds sale explainer", "Promote", "Instagram", "Story", 7),
        ("Gift wrapping tutorial", "Educate", "TikTok", "Reel", 8),
        ("Behind a wholesale order", "Behind the scenes", "YouTube",
         "Short video", 5),
        ("Ask the maker anything", "Community", "Instagram", "Story", 7),
        ("Why we price what we price", "Personal", "Instagram", "Carousel",
         8),
        ("Styling a tiny flat", "Inspire", "Pinterest", "Standard Pin", 7),
        ("New glaze test results", "Behind the scenes", "Instagram",
         "Feed post", 6),
    ]
    for i, (title, pillar, platform, fmt, score) in enumerate(idea_plan):
        used = i < 4
        s.ideas.append(dict(
            title=title, description=f"{title} - told in a {fmt.lower()}.",
            persona=["First-home Freya", "Gifting Grace",
                     "Interiors Isla"][i % 3],
            problem="People do not know how to choose or care for pieces",
            objective=OBJECTIVES.get(pillar, "Awareness"),
            pillar=pillar, platform=platform, format=fmt,
            hook=f"What nobody tells you about {title.lower()}",
            message="Handmade is for using, not for display",
            cta=CTAS.get(pillar, "Follow for more"),
            source=["Customer question", "Competitor post", "Trend",
                    "Own idea"][i % 4],
            url="", competitor="Clay & Co" if i % 5 == 0 else "",
            trend="Shelfie tour" if i % 7 == 0 else "",
            evergreen="Yes" if pillar in ("Educate", "Inspire") else "No",
            seasonal="Yes" if "gift" in title.lower() else "No",
            score=score,
            status="Used" if used else ("Shortlisted" if score >= 8
                                        else "New"),
            created=d(today, -60 + i * 3),
            reviewed=d(today, -30 + i * 2) if i % 2 else None,
            planned=d(today, 7 + i * 2) if not used else None,
            content=f"C-{i + 1:04d}" if used else "",
            notes=""))

    return s
