"""`Brand Voice` - how the brand sounds, looks and behaves online."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from

NAME = "Brand Voice"
LAST_COL = 9

BLOCKS = [
    ("Who you are", [
        ("Mission - why the business exists",
         "We make everyday objects worth keeping."),
        ("Vision - where it is going",
         "A handmade piece on every table, not in every cabinet."),
        ("Brand promise - what people can count on",
         "Made slowly, built to be used, replaced if it fails."),
        ("Brand story in one paragraph",
         "Started at a night class, kept going because people kept "
         "asking. Everything is thrown, glazed and packed by hand."),
        ("What makes you different",
         "Small batches, named glazes, and the maker answers the "
         "messages."),
    ]),
    ("How you sound", [
        ("Three words that describe your voice", "Warm, honest, playful"),
        ("Three words you are not", "Corporate, pushy, precious"),
        ("Reading age / tone", "Chatty but clear - write like you talk"),
        ("Point of view", "First person: I made this, we packed this"),
        ("Humour", "Dry and self-deprecating, never at anyone's expense"),
        ("Emoji use", "One or two, never a wall"),
        ("Exclamation marks", "One per post at most"),
    ]),
    ("Words", [
        ("Words and phrases you use",
         "handmade, small batch, slow made, everyday, kiln, glaze"),
        ("Words and phrases you avoid",
         "cheap, luxury, must-have, hurry, limited time only"),
        ("How you describe the product",
         "Stoneware made in small batches, dishwasher and microwave "
         "safe."),
        ("How you describe the customer",
         "People who use the good stuff on a Tuesday."),
        ("Your sign-off / signature line", "Made slowly, used daily."),
    ]),
    ("Rules", [
        ("Always do", "Reply to every comment in the first hour"),
        ("Never do", "Argue in the comments, or repost without asking"),
        ("How you handle a complaint",
         "Reply once publicly, move it to messages, fix it, follow up"),
        ("How you handle a troll", "Hide, do not engage, block if repeated"),
        ("How you handle a mistake",
         "Say it plainly, say what you are doing about it"),
    ]),
    ("How you look", [
        ("Logo file / link", "drive/wildflower/logo"),
        ("Colour palette", "Clay #C8734A, Sand #EAE0D2, Ink #241B38"),
        ("Fonts", "Headlines: Fraunces. Body: Inter."),
        ("Photo style",
         "Daylight, linen, slight mess - real tables, not studios"),
        ("Video style", "Hands, process, no voiceover, captions always"),
        ("Grid / feed style", "Alternating product, process, person"),
        ("Template files", "canva.com/wildflower-templates"),
    ]),
    ("Profile copy", [
        ("Bio - short (150 characters)",
         "Handmade stoneware for everyday use. Made slowly in the "
         "studio. Shop below."),
        ("Bio - long",
         "Hand-thrown ceramics made in small batches. Dishwasher safe, "
         "built to be used, shipped carefully. New glazes monthly."),
        ("Link in bio destination", "shop.wildflowerstudio.com"),
        ("Pinned post / highlight",
         "Start here: how a mug is made, and how to choose a size"),
        ("Signature hashtag", "#wildflowerstudio"),
    ]),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)

    s.banner("Brand Voice & Identity",
             "Fill this in once. Every caption you write, every reply you "
             "send and anyone you hire should match what is on this page.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 26)
    s.set_col(3, LAST_COL, 16)

    row = 4
    for title, fields in BLOCKS:
        row = s.section(row, title, LAST_COL)
        for label, example in fields:
            value = example if cfg.include_sample_data else ""
            row = s.field(row, label, value, label_col=1, value_col=3,
                          value_end=LAST_COL, style="wrap")
        row += 1

    row = s.section(row, "Tone by situation", LAST_COL)
    s.table_head(row, 1, ["Situation", "How you sound", "Example line"],
                 [26, 20, 44])
    situations = [
        ("Launching something", "Confident, warm",
         "It is here, and there are only forty."),
        ("Teaching something", "Patient, plain",
         "Here is how to tell a good glaze from a lazy one."),
        ("Sharing a mistake", "Honest, light",
         "Forty mugs, one wobbly kiln shelf. We live and learn."),
        ("Celebrating a customer", "Generous, specific",
         "Look what Hannah did with the breakfast set."),
        ("Asking for the sale", "Direct, no apology",
         "The restock is live. Link in bio."),
        ("Replying to a complaint", "Calm, quick, human",
         "That should not have happened - I am sorting it today."),
    ]
    first = row + 1
    for i, (situation, tone, example) in enumerate(situations):
        r = first + i
        value_tone = tone if cfg.include_sample_data else ""
        value_example = example if cfg.include_sample_data else ""
        s.write(r, 1, situation, fmts.chrome("label_bold"))
        s.write(r, 2, value_tone, fmts.input("text"))
        s.merge(r, 3, LAST_COL, None, fmts.input("wrap"))
        s.write(r, 3, value_example, fmts.input("wrap"))
    row = first + len(situations) + 1

    row = s.section(row, "Consistency check", LAST_COL)
    s.table_head(row, 1, ["Check", "Your answer"], [40, 20])
    checks = [
        ("Does every profile use the same name and handle?",
         "YesNoList"),
        ("Do all five profile pictures match?", "YesNoList"),
        ("Is the bio the same tone everywhere?", "YesNoList"),
        ("Do you use the same colours and fonts on graphics?",
         "YesNoList"),
        ("Does your signature hashtag appear on every post?",
         "YesNoList"),
        ("Could someone else write a post that sounds like you?",
         "YesNoList"),
    ]
    first = row + 1
    for i, (check, list_name) in enumerate(checks):
        r = first + i
        s.merge(r, 1, 2, check, fmts.chrome("label"))
        s.write(r, 3, "", fmts.input("flag"))
        s.ws.data_validation(r, 3, r, 3, list_from(list_name))
        s.ws.conditional_format(r, 3, r, 3, {
            "type": "text", "criteria": "containing", "value": "No",
            "format": fmts.cf("warn")})
        s.ws.conditional_format(r, 3, r, 3, {
            "type": "text", "criteria": "containing", "value": "Yes",
            "format": fmts.cf("good")})
    row = first + len(checks) + 1

    s.note(row,
           "A brand people recognise is worth more than a brand people "
           "like. Pick a voice you can keep up on a bad week, write it "
           "down here, and stick to it.",
           LAST_COL, first_col=1, height=28)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
