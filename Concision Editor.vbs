Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub HighlightWords()

Dim vFindMisusedWords As Variant
Dim vFindPronouns As Variant
Dim vFindFillerWords As Variant
Dim vNotAdverbs As Variant
Dim vFindText As Variant
Dim vReplText As Variant
Dim Colors As Variant
Dim i As Long
Dim j As Long

' Limits:
' if there are words with -ing or -ed, you need to add those forms of it.

'highlight possible misused words
vFindMisusedWords = Array("whose", "who's", "accept", "except", "affect", "affected", _
"effect", "effecting", "effected", "allusion", "illusion", "capital", "capitol", "emigrate", _
"emigrating", "emigrated", "immigrate", "immigrating", "immigrated", "principle", "principal", _
"two", "lie", "lay", "laying", "lying", "set", "sit", "towards", "anyways", "could care less", _
"literally")

' Pronouns
vFindPronouns = Array("he", "she", "it", "they", "we", "you", "i", "this", "that")

'highlight "Filler/Qualifying Words"
vFindFillerWords = Array("very", "several", "some", "many", "most", "few", "vast", "just", _
"quite", "often", "various", "really", "so that", "then", "next", "subsequently", _
"finally", "last", "but", "possibly", "perhaps", "maybe", "just", "sort of", "kind", "currently", _
"i consider", "i believe", "i don't believe", "i don't consider", "i don't feel", "i don't suggest", _
"i don't think", "i feel", "i hope to", "i think", "i was wondering", "i will try", "i wonder", _
"in my opinion", "we believe", "we consider", "we don't believe", "we don't consider", _
"we don't feel", "we don't suggest", "we don't think", "we feel", "we hope to", "we suggest", _
"we think", "we were wondering", "we will try", "we wonder", "opinion", "might", "in my view", _
"in our view", "in her view", "in his view", "on the other", "in fact", "later", "overall", _
"successfully")


' -----------------
'highlight Adverbs
vNotAdverbs = Array("ally", "anomaly", "fly", "barfly", "blowfly", "botfly", "mayfly", _
"medfly", "dally", "outfly", "ply", "authorly", "beastly", "brotherly", "cowardly", _
"fatherly", "gentlemanly", "granddaughterly", "housekeeperly", "husbandly", "kingly", _
"landlordly", "manly", "marksmanly", "matronly", "miserly", "motherly", "neighborly", _
"queenly", "saintly", "scholarly", "southerly", "wifely", "womanly", "easterly", _
"northeasterly", "northerly", "northwesterly", "westerly", "crumply", "frizzly", _
"rumply", "wriggly", "bodily", "knurly", "leisurely", "mannerly", "otherworldly", _
"pimply", "scaly", "shapely", "sickly", "silly", "slatternly", "slovenly", "sly", _
"spindly", "sprightly", "squiggly", "stately", "steely", "treacly", "ungainly", _
"actually", "additionally", "allegedly", "ally", "alternatively", "apply", "approximately", _
"ashely", "ashly", "assembly", "awfully", "baily", "belly", "bely", "billy", "bradly", _
"bristly", "bubbly", "bully", "burly", "butterfly", "carly", "charly", "chilly", "comely", _
"completely", "comply", "consequently", "costly", "courtly", "crinkly", "crumbly", "cuddly", _
"curly", "daily", "dastardly", "deadly", "deathly", "definitely", "dilly", "disorderly", _
"doily", "dolly", "dragonfly", "early", "elderly", "elly", "emily", "especially", _
"exactly", "exclusively", "family", "finally", "firefly", "folly", "friendly", "frilly", _
"gadfly", "gangly", "generally", "ghastly", "giggly", "globally", "goodly", "gravelly", _
"grisly", "gully", "haily", "hally", "harly", "hardly", "heavenly", "hillbilly", "hilly", _
"holly", "holy", "homely", "homily", "horsefly", "hourly", "immediately", "instinctively", _
"imply", "italy", "jelly", "jiggly", "jilly", "jolly", "july", "karly", "kelly", "kindly", _
"lately", "likely", "lilly", "lily", "lively", "lolly", "lonely", "lovely", "lowly", _
"luckily", "mealy", "measly", "melancholy", "mentally", "molly", "monopoly", "monthly", _
"multiply", "nightly", "oily", "only", "orderly", "panoply", "particularly", "partly", _
"paully", "pearly", "pebbly", "polly", "potbelly", "presumably", "previously", "pualy", _
"quarterly", "rally", "rarely", "recently", "rely", "reply", "reportedly", "roughly", "sally", "scaly", "shapely", "shelly", "shirly", "shortly", "sickly", "silly", "sly", "smelly", "sparkly", "spindly", "spritely", "squiggly", "stately", "steely", "supply", "surly", "tally", "timely", "trolly", "ugly", "underbelly", "unfortunately", "unholy", "unlikely", "usually", "waverly", "weekly", "wholly", "willy", "wily", "wobbly", "wooly", "worldly", "wrinkly", "yearly")


' highlight complex words
Dim vComplexWords As Variant

vComplexWords = Array("a number of", "abundance", "accede to", "accelerate", "accentuate", "accompany", "accomplish", _
"accorded", "accrue", "acquiesce", "acquire", "additional", "adjacent to", "adjustment", "admissible", _
"advantageous", "adversely impact", "advise", "aforementioned", "aggregate", "all of", _
"alleviate", "allocate", "along the lines of", "already existing", "alternatively", "ameliorate", _
"anticipate", "apparent", "appreciable", "as a means of", "as of yet", "as to", "as yet", "ascertain", _
"assistance", "at this time", "attain", "attributable to", "authorize", "because of the fact that", _
"belated", "benefit from", "bestow", "by virtue of", "cease", "close proximity", "commence", _
"comply with", "concerning", "consequently", "consolidate", "constitutes", "demonstrate", "depart", _
"designate", "discontinue", "due to the fact that", "each and every", "economical", "eliminate", _
"elucidate", "employ", "endeavor", "enumerate", "equitable", "equivalent", "evaluate", "evidenced", _
"exclusively", "expedite", "expend", "expiration", "facilitate", "factual evidence", "feasible", "finalize", _
"first and foremost", "for the purpose of", "forfeit", "formulate", "honest truth", "however", "if and when", "impacted", _
"implement", "in a timely manner", "in accordance with", "in addition", "in all likelihood", _
"in an effort to", "in between", "in excess of", "in lieu of", "in light of the fact that", _
"in many cases", "in order to", "in regard to", "in some instances", "in terms of", "in the near future", _
"in the process of", "inception", "incumbent upon", "indicate", "indication", "initiate", "is applicable to", _
"is authorized to", "is responsible for", "it is essential", "magnitude", "maximum", "methodology", _
"minimize", "minimum", "modify", "monitor", "multiple", "necessitate", "nevertheless", "not certain", _
"not many", "not often", "not unless", "not unlike", "notwithstanding", "null and void", "numerous", _
"objective", "obligate", "obtain", "on the contrary", "on the other hand", "one particular", "optimum", _
"owing to the fact that", "participate", "particulars", "pass away", "pertaining to", "point in time", "portion", _
"possess", "preclude", "previously", "prior to", "prioritize", "procure", "proficiency", "provided that", _
"purchase", "put simply", "readily apparent", "refer back", "regarding", "relocate", "remainder", _
"remuneration", "require", "requirement", "reside", "residence", "retain", "satisfy", "shall", _
"should you wish", "similar to", "solicit", "span across", "strategize", "subsequent", "substantial", "sufficient", "terminate", "the month of", "therefore", "this day and age", "time period", "took advantage of", "transmit", "transpire", "until such time as", "utilization", "utilize", "validate", "various different", "whether or not", "with respect to", "with the exception of", "witnessed")

Dim vComplexWordSuggestions As Variant

vComplexWordSuggestions = Array("many, some", "enough, plenty", "allow, agree to", "speed up", "stress", "go with, with", "do", "given", _
"add, gain", "agree", "get", "more, extra", "next to", "change", "allowed, accepted", "helpful", "hurt", "tell", "remove", _
"total, add", "all", "ease, reduce", "divide", "like, as in", "existing", "or", "improve, help", _
"expect", "clear, plain", "many", "to", "yet", "on, about", "yet", "find out, learn", "help", "now", _
"meet", "because", "allow, let", "because", "late", "enjoy", "give, award", "by, under", "stop", _
"near", "begin or start", "follow", "about, on", "so", "join, merge", "is, forms, makes up", _
"prove, show", "leave, go", "choose, name", "drop, stop", "because, since", "each", "cheap", _
"cut, drop, end", "explain", "use", "try", "count", "fair", "equal", "test, check", "showed", _
"only", "hurry", "spend", "end", "ease, help", "facts, evidence", "workable", "complete, finish", _
"first", "to", "lose, give up", "plan", "truth", "but, yet", "if, when", "affected, harmed, changed", _
"install, put in place, tool", "on time", "by, under", "also, besides, too", "probably", "to", "between", _
"more than", "instead", "because", "often", "to", "about, concerning, on", "sometimes", "omit", "soon", "omit", _
"start", "must", "say, state, or show", "sign", "start", "applies to", "may", "handles", _
"must, need to", "size", "greatest, largest, most", "method", "cut", "least, smallest, small", _
"change", "check, watch, track", "many", "cause, need", "still, besides, even so", "uncertain", _
"few", "rarely", "only if", "similar, alike", "in spite of, still", "use either null or void", _
"many", "aim, goal", "bind, compel", "get", "but, so", "omit, but, so", "one", "best, greatest, most", _
"because, since", "take part", "details", "die", "about, of, on", "time, point, moment, now", "part", _
"have, own", "prevent", "before", "before", "rank, focus on", "buy, get", "skill", "if", "buy, sale", _
"omit", "clear", "refer", "about, of, on", "move", "rest", "payment", "must, need", "need, rule", _
"live", "house", "keep", "meet, please", "must, will", "if you want", "like", "ask for, request", _
"span, cross", "plan", "later, next, after, then", "large, much", "enough", "end, stop", "omit", _
"thus, so", "today", "time, period", "preyed on", "send", "happen", "until", "use", "use", "confirm", _
"various, different", "whether", "on, about", "except for", "saw, seen")


'-----------------------------------------------------------
' Adverbs are in Bright Green
Options.DefaultHighlightColorIndex = wdBrightGreen
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting

' if adverb, select and highlight the word if it ends in "ly"
vReplText = "^&"
With Selection.Find
    .Forward = True
    .Wrap = wdFindContinue
    .MatchWholeWord = False
    .MatchSuffix = False
    .MatchWildcards = True
    .MatchSoundsLike = False
    .MatchPhrase = False
    .MatchAllWordForms = False
    .Format = True
    .MatchCase = False
    .Text = "<[! ]@ly>"
    .Replacement.Text = vReplText
    .Replacement.Highlight = True
    .Execute Replace:=wdReplaceAll
End With

' now un-highlight any non-adverbs:
Options.DefaultHighlightColorIndex = wdNoHighlight
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Forward = True
    .Wrap = wdFindContinue
    .MatchWholeWord = True
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchPhrase = False
    .MatchAllWordForms = False
    .Format = True
    .MatchCase = False
    
    For i = LBound(vNotAdverbs) To UBound(vNotAdverbs)
        .Text = vNotAdverbs(i)
        .Replacement.Text = vReplText
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll
    Next i
End With

'--------------------



' Complex Words highlighted in Turquoise
For i = LBound(vComplexWords) To UBound(vComplexWords)
    Options.DefaultHighlightColorIndex = wdTurquoise
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    ' if adverb, enable selection find suffix ly
    
    vReplText = "^&"
    With Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Format = True
        .MatchCase = False
        
        ' i is an index number
        .Execute FindText:=vComplexWords(i)
        .Replacement.Text = vReplText + " (suggest replacing with: " + vComplexWordSuggestions(i) + ")"
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll

    End With
Next i

'-------------------
vFindText = Array(vFindMisusedWords, vFindPronouns, vFindFillerWords)
Colors = Array(wdYellow, wdGray25, wdPink)

' Potentially misused Words in Yellow
' Pronouns in Gray
' Filler / Qualifier Words and Phrases in Pink
For i = LBound(vFindText) To UBound(vFindText)
    Options.DefaultHighlightColorIndex = Colors(i)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    ' if adverb, enable selection find suffix ly
    
    vReplText = "^&"
    With Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Format = True
        .MatchCase = False
        
        ' i is an index number
        For j = LBound(vFindText(i)) To UBound(vFindText(i))
            .Execute FindText:=vFindText(i)(j) '.Text = vFindText(i)(j)
            .Replacement.Text = vReplText
            .Replacement.Highlight = True
            .Execute Replace:=wdReplaceAll
        Next j
    End With
Next i

End Sub





