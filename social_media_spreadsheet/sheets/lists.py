"""
The `Lists` tab: every drop down list in the workbook lives here.

Each column becomes a *dynamic* defined name (OFFSET + COUNTIF) so the user
can add or remove options and every drop down in the workbook follows along.
"""

from __future__ import annotations

from collections import OrderedDict
from typing import List, Tuple

from ..config import Config, Palette
from ..layout import (DATA_ROW, HEADER_ROW, Registry, col_letter,
                      dynamic_name, quote_sheet)
from ..writer import SheetWriter

SHEET = "Lists"
ROWS = 40

#: name -> (column header, options)
LISTS: "OrderedDict[str, Tuple[str, List[str]]]" = OrderedDict([
    ("YesNoList", ("Yes / No", ["Yes", "No"])),
    ("YesNoNAList", ("Yes / No / N/A", ["Yes", "No", "N/A"])),
    ("PlatformList", ("Platform", [
        "Instagram", "TikTok", "Pinterest", "YouTube", "Facebook",
        "LinkedIn", "X (Twitter)", "Threads", "Blog", "Email newsletter",
        "Other"])),
    ("AccountTypeList", ("Account type", [
        "Business account", "Creator account", "Personal profile",
        "Page", "Group", "Channel", "Shop"])),
    ("FormatList", ("Content format", [
        "Feed post", "Carousel", "Reel", "Story", "Short video",
        "Long video", "Live", "Photo", "Graphic", "Infographic",
        "Text post", "Poll", "Standard Pin", "Video Pin", "Idea Pin",
        "Blog post", "Newsletter", "Podcast", "UGC repost"])),
    ("ContentStatusList", ("Content status", [
        "Idea", "Brief", "In production", "Draft", "Awaiting review",
        "Revision needed", "Approved", "Scheduled", "Published",
        "Archived", "Cancelled"])),
    ("ObjectiveList", ("Objective", [
        "Awareness", "Engagement", "Traffic", "Leads", "Sales",
        "Community", "Retention", "Brand authority", "Launch support"])),
    ("FunnelStageList", ("Funnel stage", [
        "Awareness", "Consideration", "Conversion", "Retention"])),
    ("CTAList", ("Call to action", [
        "Follow for more", "Save this", "Share with a friend",
        "Comment below", "Link in bio", "Shop now", "Sign up",
        "Download the guide", "Book a call", "Watch to the end",
        "Reply to this", "Tag a friend", "None"])),
    ("HookTypeList", ("Hook type", [
        "Question", "Bold statement", "Statistic", "Story", "Problem",
        "Myth bust", "How-to", "Listicle", "Before and after", "POV",
        "Tutorial", "Behind the scenes"])),
    ("ToneList", ("Tone of voice", [
        "Friendly", "Playful", "Inspiring", "Educational", "Professional",
        "Warm", "Bold", "Calm", "Witty", "Honest"])),
    ("PriorityList", ("Priority", ["High", "Medium", "Low"])),
    ("IdeaStatusList", ("Idea status", [
        "New", "Shortlisted", "Scheduled", "Used", "On hold",
        "Archived"])),
    ("IdeaSourceList", ("Idea source", [
        "Own idea", "Customer question", "Comment", "Direct message",
        "Product review", "Competitor post", "Trend", "Keyword research",
        "Analytics", "Team suggestion", "Newsletter", "Repurposed"])),
    ("SeasonalTypeList", ("Seasonal event type", [
        "Public holiday", "Awareness day", "Industry event",
        "Seasonal sale", "Product launch", "Local event",
        "Cultural moment"])),
    ("CampaignStatusList", ("Campaign status", [
        "Planned", "Running", "Paused", "Complete", "Cancelled"])),
    ("HashtagTypeList", ("Hashtag type", [
        "Broad reach", "Niche", "Community", "Branded", "Campaign",
        "Product", "Location", "Seasonal", "Trending"])),
    ("SearchIntentList", ("Search intent", [
        "Informational", "Inspirational", "Commercial", "Navigational",
        "Transactional"])),
    ("DifficultyList", ("Keyword difficulty", ["Low", "Medium", "High"])),
    ("AssetTypeList", ("Asset type", [
        "Photo", "Video clip", "Graphic", "Canva template", "Reel cover",
        "Thumbnail", "Carousel set", "Audio", "Font", "Logo", "Mockup",
        "UGC clip", "Product shot"])),
    ("UsageRightsList", ("Usage rights", [
        "Owned", "Licensed", "Royalty free", "UGC with permission",
        "Influencer content", "Stock library", "Expired"])),
    ("ApprovalStatusList", ("Approval status", [
        "Not started", "Awaiting review", "Approved", "Revision needed",
        "Rejected"])),
    ("CollabTypeList", ("Collaboration type", [
        "Gifted product", "Paid post", "Affiliate", "Brand ambassador",
        "Giveaway", "Account takeover", "Guest content", "Event"])),
    ("OutreachStatusList", ("Outreach status", [
        "Not contacted", "Contacted", "Replied", "Negotiating", "Agreed",
        "Declined", "Complete"])),
    ("UGCPermissionList", ("UGC permission", [
        "Not requested", "Requested", "Granted", "Refused", "Expired"])),
    ("UGCSourceList", ("UGC source", [
        "Tagged post", "Story mention", "Product review", "Email",
        "Direct message", "Hashtag", "Competition entry",
        "Influencer content"])),
    ("RoleList", ("Team role", [
        "Owner", "Social media manager", "Content creator", "Copywriter",
        "Designer", "Video editor", "Photographer", "Virtual assistant",
        "Contractor", "Agency"])),
    ("TaskStatusList", ("Task status", [
        "Not started", "In progress", "Blocked", "Complete"])),
    ("GoalCategoryList", ("Goal category", [
        "Followers", "Engagement", "Reach", "Website traffic", "Leads",
        "Sales", "Revenue", "Content output", "Email list",
        "Community"])),
    ("GoalStatusList", ("Goal status", [
        "Not started", "On track", "At risk", "Behind", "Achieved",
        "Missed"])),
    ("CostCategoryList", ("Cost category", [
        "Paid advertising", "Influencer fees", "Content production",
        "Design & editing", "Software & subscriptions", "Contractor fees",
        "Props & samples", "Training", "Stock media", "Printing",
        "Other"])),
    ("AdObjectiveList", ("Ad objective", [
        "Awareness", "Traffic", "Engagement", "Video views",
        "Lead generation", "Sales", "Retargeting", "App installs"])),
    ("TestVariableList", ("Test variable", [
        "Hook", "Caption length", "Thumbnail", "Posting time", "Format",
        "Call to action", "Hashtag set", "Cover image", "Audio", "Offer",
        "First comment"])),
    ("TestStatusList", ("Test status", [
        "Planned", "Running", "Complete", "Inconclusive"])),
    ("TrendTypeList", ("Trend type", [
        "Audio", "Format", "Topic", "Challenge", "Meme", "Hashtag",
        "Platform feature"])),
    ("RepurposeStatusList", ("Repurpose status", [
        "Planned", "In progress", "Published", "Not suitable"])),
    ("FrequencyList", ("Frequency", [
        "Daily", "3 x week", "Weekly", "Fortnightly", "Monthly",
        "Quarterly", "Ad hoc"])),
    ("TimeSlotList", ("Posting slot", [
        "Early morning", "Morning", "Lunchtime", "Afternoon", "Evening",
        "Night"])),
    ("ComplianceStatusList", ("Compliance status", [
        "Not checked", "Checked", "Not required", "Issue found",
        "Fixed"])),
    ("MonthList", ("Month", [
        "All year", "January", "February", "March", "April", "May",
        "June", "July", "August", "September", "October", "November",
        "December"])),
    ("QuarterList", ("Quarter", [
        "All year", "Q1", "Q2", "Q3", "Q4"])),
    ("WeekStartList", ("Week starts on", ["Monday", "Sunday"])),
])


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    sheet = SheetWriter(book, fmts, cfg, reg, SHEET,
                        tab_color=Palette.grey_text)
    last_col = len(LISTS) - 1
    sheet.banner(
        "Customisable Drop Down Lists",
        "Edit, add or delete options below - every drop down in the "
        "workbook updates automatically. Keep each list in one unbroken "
        "block, starting in row 5.",
        last_col,
        tail_text="TIP: lists grow automatically (up to row 44). Deleting "
                  "an option does not change posts that already use it.")

    header_fmt = fmts.chrome("header")
    input_fmt = fmts.input("text")
    for idx, (name, (header, values)) in enumerate(LISTS.items()):
        sheet.set_col(idx, idx, max(16, min(26, len(header) + 8)), input_fmt)
        sheet.write(HEADER_ROW - 1, idx, header, header_fmt)
        for offset, value in enumerate(values):
            sheet.write(DATA_ROW - 1 + offset, idx, value, input_fmt)
        book.define_name(name, dynamic_name(SHEET, col_letter(idx),
                                            DATA_ROW, DATA_ROW + ROWS - 1))

    sheet.ws.set_row(HEADER_ROW - 1, 30)
    sheet.ws.freeze_panes(HEADER_ROW, 0)
    sheet.finish(fit_width=False)
    return sheet


def ref(name: str) -> str:
    """A1 reference to this list's column on the current row, e.g.

        Lists!$AH{r}

    Used by the hidden Calc sheet to mirror a list down a block of rows.
    """
    idx = list(LISTS).index(name)
    return f"{quote_sheet(SHEET)}!${col_letter(idx)}{{r}}"


def options(name: str) -> List[str]:
    """Static copy of a list (used for sample data sanity)."""
    return list(LISTS[name][1])
