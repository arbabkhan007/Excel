"""`Recipes` - the recipe database.

One row per recipe: what it is, how long it takes, who it suits and how
it scored at the table. The cost and the nutrition are **not** typed in -
they are added up from the Recipe Ingredients tab, so changing the price
of butter re-prices every recipe that uses it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, score_validation

NAME = "Recipes"

#: the Recipe Ingredients columns the roll-ups sum
_LINES = "[[Recipe Ingredients.recipe]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "RecipePrefix", digits=4, width=10,
               note="The ID the meal planner, the shopping list and the "
                    "cost tabs all refer to. Generated for you."),
        Col("name", "Recipe", 30, "text",
            note="One row per recipe. The ingredients go on the Recipe "
                 "Ingredients tab, one line each."),
        Col("category", "Category", 14, "text",
            validate=list_from("RecipeCategoryList")),
        Col("meal_slot", "Usual meal", 14, "text",
            validate=list_from("MealSlotList"),
            note="Which meal of the day it normally turns up at."),
        Col("cuisine", "Cuisine", 14, "text",
            validate=list_from("CuisineList")),
        Col("description", "What it is", 34, "wrap"),
        Col("servings", "Serves", 8, "int",
            validate=number_check(1, 100, decimal=False),
            note="How many people the quantities on the Recipe "
                 "Ingredients tab feed. Everything scales from this."),
        Col("prep_min", "Prep (min)", 10, "int",
            validate=number_check(0, 600, decimal=False)),
        Col("cook_min", "Cook (min)", 10, "int",
            validate=number_check(0, 1440, decimal=False)),
        Col("total_min", "Total (min)", 11, "int",
            formula='=IF({name}{r}="","",IF(AND({prep_min}{r}="",'
                    '{cook_min}{r}=""),"",'
                    'IF({prep_min}{r}="",0,{prep_min}{r})+'
                    'IF({cook_min}{r}="",0,{cook_min}{r})))',
            cf=[{"type": "cell", "criteria": "<=", "value": 30,
                 "fmt": "good"},
                {"type": "cell", "criteria": ">", "value": 90,
                 "fmt": "warn"}]),
        Col("quick", "Quick meal?", 11, "flag",
            formula='=IF({total_min}{r}="","",'
                    'IF({total_min}{r}<=30,"Yes","No"))',
            note="30 minutes or less, start to finish."),
        Col("method", "Method", 14, "text",
            validate=list_from("CookingMethodList")),
        Col("difficulty", "Difficulty", 11, "text",
            validate=list_from("DifficultyList")),
        Col("equipment", "Equipment", 20, "text",
            note="Anything you need out of the cupboard: blender, "
                 "slow cooker, baking tin."),
        Col("one_pot", "One pot?", 9, "flag", validate=list_from(
            "YesNoList")),
        Col("no_cook", "No cook?", 9, "flag", validate=list_from(
            "YesNoList")),
        Col("prep_friendly", "Batch / prep friendly?", 13, "flag",
            validate=list_from("YesNoList"),
            note="Does it survive being made in advance and reheated?"),
        Col("freezes", "Freezes well?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("kid_friendly", "Kid friendly?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("diet", "Suits", 15, "text", validate=list_from("DietList"),
            note="The dietary filters use this plus the allergens "
                 "worked out from the ingredients."),
        Col("spice", "Spice level", 11, "text",
            validate=list_from("SpiceLevelList")),
        Col("season", "Best season", 12, "text",
            validate=list_from("SeasonList")),
        Col("tags", "Tags", 20, "text",
            note="Free text: 'payday treat', 'uses up leftovers', "
                 "'lunchbox'. Filter on it."),
        Col("source", "Where it came from", 20, "text"),
        Col("link", "Link", 24, "link",
            note="Paste the URL, or the page number of the book."),
        Col("photo", "Photo link", 18, "link",
            note="A spreadsheet cannot keep the photo itself tidy - "
                 "paste a cloud link instead."),
        Col("method_notes", "How to make it", 46, "wrap",
            note="Short numbered steps. 1. Chop... 2. Fry... Keep it to "
                 "what you would scribble on a card."),
        # ------------------------------------------------ rolled up totals
        Col("line_count", "Ingredients", 11, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _LINES + ',{id}{r}))',
            note="Counted from the Recipe Ingredients tab."),
        Col("cost", "Cost to make", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS('
                    '[[Recipe Ingredients.line_cost]],'
                    + _LINES + ',{id}{r}))',
            note="Added up from the ingredient prices. Change a price "
                 "once and every recipe re-prices itself."),
        Col("cost_serving", "Cost per serving", 13, "money",
            formula='=IF(OR({cost}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",{cost}{r}/{servings}{r}))',
            cf=[{"type": "3_color_scale", "min_color": "#D9F2E4",
                 "mid_color": "#FFF1CF", "max_color": "#FBD9DD"}]),
        Col("calories", "Calories a serving", 13, "number",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_cals]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("protein", "Protein (g)", 11, "number1",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_protein]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("carbs", "Carbs (g)", 10, "number1",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_carbs]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("fat", "Fat (g)", 9, "number1",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_fat]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("fibre", "Fibre (g)", 10, "number1",
            formula='=IF(OR({id}{r}="",{servings}{r}=""),"",'
                    'IF({servings}{r}=0,"",SUMIFS('
                    '[[Recipe Ingredients.line_fibre]],'
                    + _LINES + ',{id}{r})/{servings}{r}))'),
        Col("allergens", "Allergens in it", 22, "text",
            formula='=IF({id}{r}="","",IF(COUNTIFS(' + _LINES + ',{id}{r},'
                    '[[Recipe Ingredients.allergen]],"<>",'
                    '[[Recipe Ingredients.allergen]],"<>None")=0,'
                    '"None listed",COUNTIFS(' + _LINES + ',{id}{r},'
                    '[[Recipe Ingredients.allergen]],"<>",'
                    '[[Recipe Ingredients.allergen]],"<>None")&'
                    '" ingredient(s) flagged"))',
            note="Counted from the allergen column of each ingredient. "
                 "Always read the actual labels as well.",
            cf=[{"type": "text", "criteria": "containing",
                 "value": "flagged", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "None listed", "fmt": "good"}]),
        Col("plan_scale", "Batches in the shopping window", 12, "number1",
            formula='=IF({id}{r}="",0,SUMIFS([[Meal Plan.scale]],'
                    '[[Meal Plan.recipe]],{id}{r},[[Meal Plan.date]],'
                    '">="&ShopFrom,[[Meal Plan.date]],"<="&ShopTo,'
                    '[[Meal Plan.shop]],"Yes"))',
            hidden=True,
            note="How many recipes-worth the grocery list has to buy "
                 "for. Driven by the dates on the Grocery List tab."),
        Col("plan_cost", "Cost in the window", 12, "money",
            formula='=IF(OR({id}{r}="",{cost}{r}=""),"",'
                    '{cost}{r}*{plan_scale}{r})',
            hidden=True),
        Col("missing_price", "Missing a price", 12, "int",
            formula='=IF({id}{r}="","",COUNTIFS(' + _LINES + ',{id}{r},'
                    '[[Recipe Ingredients.line_cost]],""))',
            hidden=True),
        # ------------------------------------------------------- the family
        Col("rating", "Your rating", 11, "rating",
            validate=score_validation(5),
            note="1 to 5. The rotation planner leans on this."),
        Col("family_score", "Family score", 11, "rating",
            validate=score_validation(5),
            note="What everyone else thought."),
        Col("favourite", "Favourite?", 10, "flag",
            validate=list_from("YesNoList")),
        Col("times_planned", "Times planned", 12, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Meal Plan.recipe]],'
                    '{id}{r}))'),
        Col("times_cooked", "Times cooked", 12, "int",
            formula='=IF({id}{r}="","",COUNTIFS([[Meal Plan.recipe]],'
                    '{id}{r},[[Meal Plan.status]],"Prepared")+'
                    'COUNTIFS([[Meal Plan.recipe]],{id}{r},'
                    '[[Meal Plan.status]],"Eaten"))'),
        Col("last_cooked", "Last cooked", 12, "date",
            formula='=IF({id}{r}="","",IF(COUNTIFS([[Meal Plan.recipe]],'
                    '{id}{r},[[Meal Plan.date]],"<="&TODAY())=0,"",'
                    'SUMPRODUCT(MAX(([[Meal Plan.recipe]]={id}{r})*'
                    '([[Meal Plan.date]]<=TODAY())*'
                    '[[Meal Plan.date]]))))',
            note="Automatic, from the meal plan."),
        Col("days_since", "Days since", 10, "int",
            formula='=IF({last_cooked}{r}="","",TODAY()-'
                    '{last_cooked}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=RotationGapDays",
                 "fmt": "warn"}]),
        Col("notes", "Tweaks & notes", 28, "wrap",
            note="'Double the garlic', 'kids pick out the peppers', "
                 "'needs 10 minutes longer in our oven'."),
        Col("added", "Added on", 11, "date", validate=date_check()),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "name",
                'IF({servings}{r}="","Add - how many does it serve?",'
                'IF({line_count}{r}=0,"Add - no ingredients listed yet",'
                'IF({missing_price}{r}>0,"Check - '
                '"&{missing_price}{r}&" ingredient(s) have no price",'
                'IF(AND({days_since}{r}<>"",'
                '{days_since}{r}<RotationGapDays),'
                '"Cooked recently","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Recipe Database",
        subtitle="Every recipe you cook, with its cost and nutrition "
                 "worked out from the ingredients rather than typed in.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).recipes,
        help_text="Add the recipe here first, then list its ingredients "
                  "on the Recipe Ingredients tab using the ID in column "
                  "A. That is what makes the grocery list automatic.",
        extra_notes=[
            "Serves matters more than anything else on this tab: the "
            "shopping list multiplies the ingredients by (people you "
            "are feeding / serves), so get it right and everything "
            "else follows.",
            "Cost, calories and allergens are grey because they are "
            "added up from the ingredient lines - there is nothing to "
            "type.",
            "Filter the Suits, Quick meal and Freezes columns to answer "
            "'what can I make in 20 minutes that the vegetarian will "
            "eat?'",
        ],
    )
