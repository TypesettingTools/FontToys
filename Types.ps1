Add-Type -TypeDefinition @"
   [System.Flags]
   public enum fsSelection
   {
      Invalid = 0,
      Italic = 1,
      Underscore = 2,
      Negative = 4,
      Outlined = 8,
      Strikeout = 16,
      Bold = 32,
      Regular = 64,
      UseTypoMetrics = 128,
      WWS = 256,
      Oblique = 512
   }
"@

Add-Type -TypeDefinition @"
   public enum usWeightClass
   {
      Thin = 100,
      ExtraLight = 200,
      Light = 300,
      Regular = 400,
      Medium = 500,
      SemiBold = 600,
      Bold = 700,
      ExtraBold = 800,
      Black = 900,
      Fat = 1000
   }
"@

Add-Type -TypeDefinition @"
    public enum usWidthClass
    {
        UltraCondensed = 1,
        ExtraCondensed = 2,
        Condensed = 3,
        SemiCondensed = 4,
        Normal = 5,
        SemiExpanded = 6,
        Expanded = 7,
        ExtraExpanded = 8,
        UltraExpanded = 9
    }
"@