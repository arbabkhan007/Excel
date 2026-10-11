"""The five platform planners, built from one shared factory.

Every planner works the same way: pick a content ID from the Content
Calendar, fill in the fields that only matter on that platform, and let the
grey columns pull the title, date, status and the shared results (reach,
engagement, clicks) from the Performance tab.  Nothing is ever typed twice.
"""

from __future__ import annotations

from typing import Any, Dict, List

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, bar_cf, id_col, lookup_content, rate_cf, yesno

CALENDAR = "Content Calendar"


# --------------------------------------------------------------------- base

def _head(platform: str, prefix: str) -> List[Col]:
    """Content ID + the details pulled back from the Content Calendar."""
    return [
        id_col("content", f'"{prefix}"', digits=3, width=9),
        Col("content", "Content ID", 12, "text",
            validate=list_from("ContentIDs"),
            note=f"Pick the C-0000 ID of the {platform} post from the "
                 f"Content Calendar. Everything grey fills itself in."),
        Col("title", "Content title", 30, "text",
            formula=lookup_content("title")),
        Col("date", "Publish date", 12, "date",
            formula=lookup_content("date")),
        Col("pillar", "Pillar", 15, "text",
            formula=lookup_content("pillar")),
        Col("status", "Status", 14, "status",
            formula=lookup_content("status")),
    ]


def _shared_results() -> List[Col]:
    """Results every platform shares - pulled from the Performance tab."""
    pull = ('=IF({content}{r}="","",IF(COUNTIF([[Performance.content]],'
            '{content}{r})=0,"",SUMIFS([[Performance.%s]],'
            '[[Performance.content]],{content}{r})))')
    return [
        Col("reach", "Reach", 11, "int", formula=pull % "reach",
            note="From the Performance tab - enter each post's numbers "
                 "once, there.",
            cf=bar_cf()),
        Col("impressions", "Impressions", 11, "int",
            formula=pull % "impressions"),
        Col("engagements", "Engagements", 11, "int",
            formula=pull % "engagements"),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("saves", "Saves", 9, "int", formula=pull % "saves"),
        Col("shares", "Shares", 9, "int", formula=pull % "shares"),
        Col("comments", "Comments", 10, "int", formula=pull % "comments"),
        Col("clicks", "Clicks", 9, "int", formula=pull % "clicks"),
    ]


def _tail(alert_body: str) -> List[Col]:
    return [
        Col("alert", "Action needed", 28, "status",
            formula=f'=IF({{content}}{{r}}="","",{alert_body})',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]


class Planner:
    """A tiny module-like object so `build.py` can treat it as a sheet."""

    def __init__(self, name: str, platform: str, prefix: str, title: str,
                 subtitle: str, help_text: str, middle, alert: str,
                 samples: str, tab_color: str, rows_key: str) -> None:
        self.NAME = name
        self.platform = platform
        self.prefix = prefix
        self.title = title
        self.subtitle = subtitle
        self.help_text = help_text
        self.middle = middle
        self.alert = alert
        self.samples = samples
        self.tab_color = tab_color
        self.rows_key = rows_key

    def spec(self, cfg: Config) -> SheetSpec:
        columns = (_head(self.platform, self.prefix) + self.middle()
                   + _shared_results() + _tail(self.alert))
        return SheetSpec(
            name=self.NAME,
            title=self.title,
            subtitle=self.subtitle,
            columns=columns,
            rows=cfg.row_count(self.rows_key),
            tab_color=self.tab_color,
            freeze_cols=3,
            samples=getattr(sample(cfg), self.samples),
            help_text=self.help_text,
        )


def _secs(lo: int, hi: int) -> Dict[str, Any]:
    return number_check(lo, hi, decimal=False)


# ---------------------------------------------------------------- instagram

def _instagram_cols() -> List[Col]:
    return [
        Col("format", "Instagram format", 14, "text",
            validate=list_from("FormatList")),
        Col("hook", "Hook / first slide", 32, "wrap"),
        Col("slides", "Carousel slides", 11, "int",
            validate=_secs(0, 20),
            note="Carousels of 6-10 slides tend to hold attention best."),
        Col("duration", "Reel length (sec)", 12, "int",
            validate=_secs(0, 900)),
        Col("cover", "Cover image / file", 22, "text"),
        Col("story_sequence", "Story sequence", 28, "wrap",
            note="Plan the frames: hook, proof, poll, link."),
        Col("collab", "Collab post?", 11, "flag", validate=yesno()),
        Col("profile_visits", "Profile visits", 12, "int",
            validate=_secs(0, 10000000)),
        Col("followers_gained", "Followers gained", 13, "int",
            validate=_secs(-100000, 10000000)),
    ]


INSTAGRAM = Planner(
    name="Instagram", platform="Instagram", prefix="IG-",
    title="Instagram Planner",
    subtitle="Grid, reels and stories in one place. Fill in the Instagram "
             "specific fields - the title, date and results come from the "
             "calendar and the Performance tab.",
    help_text="Reels under 30 seconds and carousels of 6-10 slides are the "
              "usual sweet spots. Plan your story frames in the sequence "
              "column.",
    middle=_instagram_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF(AND({format}{r}="Reel",{duration}{r}=0),'
          '"Add - reel length missing",IF({cover}{r}="",'
          '"Add - no cover image",IF(AND({reach}{r}<>"",'
          '{er}{r}<TargetEngagement),"Review - below your target rate",'
          '"OK"))))',
    samples="instagram", tab_color=Palette.primary_mid,
    rows_key="Instagram")


# ------------------------------------------------------------------- tiktok

def _tiktok_cols() -> List[Col]:
    return [
        Col("hook", "First 2 seconds", 30, "wrap",
            note="TikTok lives or dies on the first two seconds."),
        Col("sound", "Sound / audio", 24, "text"),
        Col("script", "Script outline", 32, "wrap"),
        Col("shots", "Shot list", 28, "wrap"),
        Col("duration", "Length (sec)", 11, "int", validate=_secs(0, 1800)),
        Col("text_on_screen", "Text on screen", 24, "wrap"),
        Col("cover", "Cover frame", 20, "text"),
        Col("series", "Series name", 18, "text"),
        Col("watch_time", "Avg watch time (sec)", 13, "number1",
            validate=number_check(0, 100000)),
        Col("completion", "Completion rate", 12, "percent",
            validate=number_check(0, 1),
            cf=[{"type": "cell", "criteria": ">=", "value": 0.5,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.25,
                 "fmt": "warn"}]),
        Col("profile_views", "Profile views", 11, "int",
            validate=_secs(0, 10000000)),
        Col("followers_gained", "Followers gained", 13, "int",
            validate=_secs(-100000, 10000000)),
    ]


TIKTOK = Planner(
    name="TikTok", platform="TikTok", prefix="TT-",
    title="TikTok Planner",
    subtitle="Hooks, sounds, scripts and shot lists, plus the watch time "
             "and completion rate that decide how far a video travels.",
    help_text="Write the hook as the words on screen in the first two "
              "seconds. Completion rate matters more than length.",
    middle=_tiktok_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF({hook}{r}="","Add - no hook written",'
          'IF({sound}{r}="","Add - no sound chosen",'
          'IF(AND({completion}{r}<>"",{completion}{r}<0.25),'
          '"Review - most people scroll away early","OK"))))',
    samples="tiktok", tab_color=Palette.primary_mid, rows_key="TikTok")


# ---------------------------------------------------------------- pinterest

def _pinterest_cols() -> List[Col]:
    return [
        Col("board", "Board", 22, "text"),
        Col("board_url", "Board link", 26, "link"),
        Col("pin_title", "Pin title", 30, "text",
            note="Pinterest is a search engine - put the keyword in the "
                 "title."),
        Col("pin_description", "Pin description", 38, "wrap"),
        Col("destination", "Destination link", 28, "link"),
        Col("image", "Image file", 20, "text"),
        Col("dimensions", "Size (2:3 is best)", 13, "text"),
        Col("keyword", "Target keyword", 24, "text"),
        Col("intent", "Search intent", 14, "text",
            validate=list_from("SearchIntentList")),
        Col("seasonal", "Seasonal pin?", 11, "flag", validate=yesno(),
            note="Seasonal pins want publishing 30-45 days early."),
        Col("outbound_clicks", "Outbound clicks", 12, "int",
            validate=_secs(0, 10000000)),
        Col("audience_growth", "Audience growth", 12, "int",
            validate=_secs(-100000, 10000000)),
        Col("refresh_due", "Refresh / repin on", 13, "date"),
    ]


PINTEREST = Planner(
    name="Pinterest", platform="Pinterest", prefix="PN-",
    title="Pinterest Planner",
    subtitle="Pins keep working for months, so plan the keyword, the board "
             "and the refresh date as carefully as the image.",
    help_text="Use 2:3 images (1000 x 1500), put your keyword in the pin "
              "title and description, and publish seasonal pins 30-45 days "
              "ahead.",
    middle=_pinterest_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF({keyword}{r}="","Add - no target keyword",'
          'IF({destination}{r}="","Add - no destination link",'
          'IF(AND({refresh_due}{r}<>"",{refresh_due}{r}<TODAY()),'
          '"Due - time to refresh this pin","OK"))))',
    samples="pinterest", tab_color=Palette.primary_mid,
    rows_key="Pinterest")


# ------------------------------------------------------------------ youtube

def _youtube_cols() -> List[Col]:
    return [
        Col("format", "Long or Short", 12, "text",
            validate=list_from("FormatList")),
        Col("title_option", "Video title", 32, "text",
            note="Under 60 characters so it is not cut off in search."),
        Col("title_chars", "Title length", 10, "int",
            formula='=IF({title_option}{r}="","",LEN({title_option}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 60,
                 "fmt": "warn"}]),
        Col("hook", "First 15 seconds", 30, "wrap"),
        Col("description", "Description", 36, "wrap"),
        Col("script", "Script / structure", 32, "wrap"),
        Col("chapters", "Chapters", 26, "wrap"),
        Col("thumbnail", "Thumbnail file", 20, "text"),
        Col("thumbnail_idea", "Thumbnail idea", 26, "wrap"),
        Col("keywords", "Target keywords", 26, "wrap"),
        Col("tags", "Tags", 24, "wrap"),
        Col("playlist", "Playlist", 18, "text"),
        Col("end_screen", "End screen / cards", 22, "text"),
        Col("watch_time", "Watch time (min)", 13, "number1",
            validate=number_check(0, 10000000)),
        Col("avg_duration", "Avg view (min)", 12, "number1",
            validate=number_check(0, 100000)),
        Col("retention", "Audience retention", 12, "percent",
            validate=number_check(0, 1),
            cf=[{"type": "cell", "criteria": ">=", "value": 0.5,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.3,
                 "fmt": "warn"}]),
        Col("ctr", "Thumbnail CTR", 11, "percent",
            validate=number_check(0, 1),
            cf=[{"type": "cell", "criteria": ">=", "value": 0.05,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.03,
                 "fmt": "warn"}]),
        Col("subs_gained", "Subscribers gained", 13, "int",
            validate=_secs(0, 10000000)),
        Col("subs_lost", "Subscribers lost", 12, "int",
            validate=_secs(0, 10000000)),
        Col("net_subs", "Net subscribers", 12, "int",
            formula='=IF({subs_gained}{r}="","",{subs_gained}{r}'
                    '-N({subs_lost}{r}))'),
        Col("traffic", "Main traffic source", 20, "text"),
    ]


YOUTUBE = Planner(
    name="YouTube", platform="YouTube", prefix="YT-",
    title="YouTube Planner",
    subtitle="Titles, thumbnails, chapters and the two numbers that drive "
             "the algorithm: click through rate and retention.",
    help_text="A 4-6% thumbnail CTR and 50%+ retention are healthy. Keep "
              "titles under 60 characters and always fill in the chapters.",
    middle=_youtube_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF({thumbnail}{r}="","Add - no thumbnail",'
          'IF(AND({ctr}{r}<>"",{ctr}{r}<0.03),'
          '"Review - thumbnail click rate is low",'
          'IF(AND({retention}{r}<>"",{retention}{r}<0.3),'
          '"Review - viewers leave early","OK"))))',
    samples="youtube", tab_color=Palette.primary_mid, rows_key="YouTube")


# ----------------------------------------------------------------- facebook

def _facebook_cols() -> List[Col]:
    return [
        Col("format", "Post format", 14, "text",
            validate=list_from("FormatList")),
        Col("audience", "Audience", 22, "text"),
        Col("targeting", "Targeting notes", 26, "wrap"),
        Col("link", "Link shared", 26, "link"),
        Col("group_post", "Posted to a group?", 12, "flag",
            validate=yesno()),
        Col("event", "Linked event", 20, "text"),
        Col("reactions", "Reactions", 11, "int",
            validate=_secs(0, 10000000)),
        Col("link_clicks", "Link clicks", 11, "int",
            validate=_secs(0, 10000000)),
        Col("followers_gained", "Page followers gained", 14, "int",
            validate=_secs(-100000, 10000000)),
    ]


FACEBOOK = Planner(
    name="Facebook", platform="Facebook", prefix="FB-",
    title="Facebook Planner",
    subtitle="Page posts, groups and events - the place where links still "
             "work and your existing customers still hang out.",
    help_text="Facebook rewards conversation. Reply to every comment in "
              "the first hour and keep link posts short.",
    middle=_facebook_cols,
    alert='IF({title}{r}="Not found","Check - content ID not on the '
          'calendar",IF(AND({link}{r}="",{format}{r}<>"Photo"),'
          '"Add - no link on a link post",IF(AND({reach}{r}<>"",'
          '{er}{r}<TargetEngagement),"Review - below your target rate",'
          '"OK")))',
    samples="facebook", tab_color=Palette.primary_mid, rows_key="Facebook")


#: every planner, in tab order
PLANNERS = [INSTAGRAM, TIKTOK, PINTEREST, YOUTUBE, FACEBOOK]
