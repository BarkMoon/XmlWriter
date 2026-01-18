// TestCode Generated From Antigravity - Gemini 3 Pro
// Generated Script from @TableName
// Date: @GeneratedDate

using System;
using System.Collections.Generic;

public class @TableName_Data
{
    public static void LoadData()
    {
#ForAllData
        // ID: ${Id}
        var item${Id} = new Item();
        item${Id}.Suit = "${Properties.Suit}";
        item${Id}.Number = ${Status.Number};
        #If(#Eq(${Properties.Suit}, Spade))
        item${Id}.IsTrumps = true;
        #Endif
        
        Dictionary.Add(${Id}, item${Id});

#EndForAllData
    }
}
