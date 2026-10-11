"""
The `Settings` tab - business setup, brand basics, targets and ID prefixes.

Everything the rest of the workbook needs to know about the brand is entered
once here and published as a defined name (BrandName, PlanYear,
TargetEngagement, ProductionDays ...) so the formulas stay readable.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

SHEET = "Settings"
LAST_COL = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    b = cfg.brand

    s.banner("Business Setup & Settings",
             "Fill in the cream cells once. Grey cells are locked formulas. "
             "Everything else in the workbook reads these values.",
             LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 26, fmts.chrome("label"))
    s.set_col(2, 2, 16, fmts.chrome("label"))
    s.set_col(3, 3, 30)
    s.set_col(4, 4, 20)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5

    # ------------------------------------------------------------ 1. brand
    r = s.section(r, "1. Brand details", LAST_COL)
    r = s.field(r, "Business or brand name", b.name, name="BrandName",
                value_end=4,
                help_text="Shown on every printable page and report.")
    r = s.field(r, "Owner / main contact", b.owner, name="BrandOwner",
                value_end=4)
    r = s.field(r, "Email", b.email, name="BrandEmail", value_end=4)
    r = s.field(r, "Phone", b.phone, name="BrandPhone", value_end=4)
    r = s.field(r, "Website", b.website, name="BrandWeb", value_end=4)
    r = s.field(r, "Shop / booking link", b.shop, name="BrandShop",
                value_end=4,
                help_text="The link you send people to - used in the "
                          "campaign and launch planners.")
    r = s.field(r, "Industry", b.industry, name="BrandIndustry",
                value_end=4)
    r = s.field(r, "Niche / speciality", b.niche, name="BrandNiche",
                value_end=4)
    r = s.field(r, "Target audience", b.audience, name="BrandAudience",
                value_end=4,
                help_text="One sentence. The Audience Personas tab breaks "
                          "this down properly.")
    r = s.field(r, "Main marketing goal", b.goal, name="BrandGoal",
                value_end=4)
    r = s.field(r, "Branded hashtag", b.hashtag, name="BrandHashtag",
                value_end=4)

    # -------------------------------------------------------- 2. identity
    r += 1
    r = s.section(r, "2. Visual identity reference", LAST_COL)
    r = s.field(r, "Logo file / link", b.logo_link, name="BrandLogo",
                value_end=4,
                help_text="Paste a cloud link. Spreadsheets cannot preview "
                          "an external image, so keep the link handy.")
    r = s.field(r, "Brand colour palette", b.palette, name="BrandPalette",
                value_end=4)
    r = s.field(r, "Fonts / typography", b.fonts, name="BrandFonts",
                value_end=4)
    r = s.field(r, "Brand asset folder", "brand-assets/", name="BrandFolder",
                value_end=4)

    # -------------------------------------------------------- 3. calendar
    r += 1
    r = s.section(r, "3. Calendar & time settings", LAST_COL)
    r = s.field(r, "Planning year", cfg.year, style="int", name="PlanYear",
                validate=number_check(2000, 2100, decimal=False),
                help_text="The year the dashboards and the 12 month "
                          "calendar open on.")
    r = s.field(r, "Week starts on", cfg.week_start, name="WeekStart",
                validate=list_from("WeekStartList"),
                help_text="Used by the week numbers on the calendar.")
    r = s.field(r, "Time zone", b.timezone, name="BrandTimeZone",
                value_end=4,
                help_text="Posting times on the calendar are in this zone.")
    r = s.field(r, "Currency code", cfg.currency_code, name="CurrencyCode",
                help_text="Budget and revenue columns use the currency "
                          "symbol set when the workbook was built.")
    r = s.field(r, "Fiscal year starts in month", 1, style="int",
                name="FiscalStartMonth",
                validate=number_check(1, 12, decimal=False),
                help_text="1 = January. Used by the quarterly report.")

    # --------------------------------------------------------- 4. targets
    r += 1
    r = s.section(r, "4. Targets (the dashboards measure against these)",
                  LAST_COL)
    r = s.field(r, "Posts per week target", cfg.posts_per_week_target,
                style="int", name="PostsPerWeek",
                validate=number_check(0, 500, decimal=False))
    r = s.field(r, "Posts per month target", cfg.monthly_post_target,
                style="int", name="MonthlyPostTarget",
                validate=number_check(0, 2000, decimal=False),
                help_text="Actual against planned is on the dashboard.")
    r = s.field(r, "Engagement rate target", cfg.engagement_target,
                style="percent", name="TargetEngagement",
                validate=number_check(0, 1),
                help_text="Posts at or above this rate turn green.")
    r = s.field(r, "Follower goal (end of year)", cfg.follower_goal,
                style="int", name="FollowerGoal",
                validate=number_check(0, 100000000, decimal=False))
    r = s.field(r, "Reach goal (per month)", cfg.reach_goal, style="int",
                name="ReachGoal",
                validate=number_check(0, 100000000, decimal=False))
    r = s.field(r, "Website click goal (per month)", cfg.click_goal,
                style="int", name="ClickGoal",
                validate=number_check(0, 10000000, decimal=False))
    r = s.field(r, "Lead goal (per month)", cfg.lead_goal, style="int",
                name="LeadGoal",
                validate=number_check(0, 1000000, decimal=False))
    r = s.field(r, "Conversion goal (per month)", cfg.conversion_goal,
                style="int", name="ConversionGoal",
                validate=number_check(0, 1000000, decimal=False))
    r = s.field(r, "Revenue goal (per year)", cfg.revenue_goal,
                style="money", name="RevenueGoal",
                validate=number_check(0, 100000000))

    # ---------------------------------------------------------- 5. budget
    r += 1
    r = s.section(r, "5. Budget & return", LAST_COL)
    r = s.field(r, "Marketing budget per month", cfg.monthly_budget,
                style="money", name="MonthlyBudget",
                validate=number_check(0, 10000000))
    r = s.field(r, "Marketing budget per year", cfg.annual_budget,
                style="money", name="AnnualBudget",
                validate=number_check(0, 100000000))
    r = s.field(r, "Return on ad spend target", cfg.roas_target,
                style="number1", name="RoasTarget",
                validate=number_check(0, 1000),
                help_text="4 means every 1 spent should bring back 4.")
    r = s.field(r, "Cost per lead target", cfg.cost_per_lead_target,
                style="money", name="CplTarget",
                validate=number_check(0, 100000))

    # -------------------------------------------------------- 6. workflow
    r += 1
    r = s.section(r, "6. Workflow & reminder windows (days)", LAST_COL)
    r = s.field(r, "Draft ready before publishing", cfg.production_days,
                style="int", name="ProductionDays",
                validate=number_check(0, 90, decimal=False),
                help_text="The calendar works out your draft deadline from "
                          "the publishing date.")
    r = s.field(r, "Approval needed before publishing", cfg.approval_days,
                style="int", name="ApprovalDays",
                validate=number_check(0, 90, decimal=False))
    r = s.field(r, "Idea review window", cfg.idea_review_days, style="int",
                name="IdeaReviewDays",
                validate=number_check(0, 365, decimal=False),
                help_text="Ideas older than this still sitting at 'New' are "
                          "flagged.")
    r = s.field(r, "Trend acts-fast window", cfg.trend_window_days,
                style="int", name="TrendWindowDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Asset refresh window", cfg.asset_review_days,
                style="int", name="AssetReviewDays",
                validate=number_check(0, 3650, decimal=False))
    r = s.field(r, "Account counts as quiet after", cfg.inactive_account_days,
                style="int", name="QuietAccountDays",
                validate=number_check(0, 365, decimal=False))

    # ------------------------------------------------------- 7. ID prefix
    r += 1
    r = s.section(r, "7. Reference prefixes", LAST_COL)
    r = s.field(r, "Content ID prefix", "C-", name="ContentPrefix",
                help_text="C-0001, C-0002 ... on the Content Calendar.")
    r = s.field(r, "Idea ID prefix", "IDEA-", name="IdeaPrefix")
    r = s.field(r, "Caption ID prefix", "CAP-", name="CaptionPrefix")
    r = s.field(r, "Hashtag ID prefix", "TAG-", name="HashtagPrefix")
    r = s.field(r, "Campaign ID prefix", "CMP-", name="CampaignPrefix")
    r = s.field(r, "Asset ID prefix", "AS-", name="AssetPrefix")

    r += 1
    r = s.note(r, "Tip: everything on this tab is a named cell, so you can "
                  "use the names in your own formulas too - type "
                  "=TargetEngagement or =BrandName in any empty cell to see. "
                  "The drop down options live on the Lists tab.",
               LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
