Attribute VB_Name( = "NameFunctions")
 function WorkbookName(Branch  , mp1  , mp2  , sourceRte1  , sourceRte2  )  {
    if( Branch == "" ) {
        WorkbookName( = "")
         function{
     }
    var suffix  
    
    if( Left(Right(Branch, FindInStr(Branch, "_") + 1), 1) == "R" ) {
        Branch( = Left(Branch, FindInStr(Branch, "_") - 1))
     }
    if( sourceRte1 != "" ) {
        if( sourceRte2 != "" ) {
            suffix = "_R" + sourceRte1 + "_" + sourceRte2
        }else{
            suffix = "_R" + sourceRte1
         }
    }else{
        if( sourceRte2 != "" ) {
            suffix = "_R" + sourceRte2
         }
     }
    if( Left(Branch, 1) == "U" ) {
        WorkbookName = Branch + "_" + Format(Now, "DDMMMYYYY") + suffix + ".xlsm"
    }else{
        WorkbookName = Branch + "_MP" + Format(mp1, "0.0000") + "-" + Format(mp2, "0.0000") + "_" + Format(Now, "DDMMMYYYY") + suffix + ".xlsm"
     }
 }
 function UnnamedBranch(branchType  , srcRte  )  {
    var rteStr  
    var sufRte  
    if( branchType == "" ) {
        if( srcRte == "" ) {
            UnnamedBranch( = "")
             function{
            
        }else{
            branchType( = "DREG")
         }
     }
    //does not handle multiple source Routes appended to the suffix
    //if the srcRte is Unnamed, handle with care
    //if the srcrte is blank, nothing
    //if the srtrte is not blank, not unnamed then add "_R"
    //if unnamed srcrte, check for its srcrte suffix and pull it off
    //   then name it as srcrte + u + new timestamp + the og suffix
    rteStr( = "")
    if( Left(srcRte, 2) == "U_" ) {
        sufRte( = Right(srcRte, Len(srcRte) - FindInStr(srcRte, "_")))
        if( Left(sufRte, 1) == "R" ) {
            srcRte( = Left(srcRte, FindInStr(srcRte, "_") - 1))
            UnnamedBranch = srcRte + "_U" + TimeStamp + "_" + sufRte
             function{
         }
        UnnamedBranch = srcRte + "_U" + TimeStamp
    }else{
        if( ! srcRte == "" ) {
            rteStr = "_R" + srcRte
         }
        UnnamedBranch = "U_" + branchType + "_" + TimeStamp + rteStr
     }
 }
 function TimeStamp()  {
    TimeStamp( = Format(Now(), "YYYYMMDDHHNN"))
 }
 function FindInStr(FindIn  , ToFind  )  Integer{
    //searches for a character starting from the end of a string
    var FindCha  Integer
    for( FindCha = Len(FindIn) - Len(ToFind) + 1 To 1 Step -1) {
        if( Mid(FindIn, FindCha, Len(ToFind)) == ToFind ) {
            FindInStr( = FindCha)
             function{
         }
    } FindCha
 }


