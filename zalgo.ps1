function New-Zalgo {

    #This implementation in PowerShell is based off gregoryneal's implementation in python here:
    #https://github.com/gregoryneal/zalgo/blob/master/code/zalgo_text/zalgo.py

    param(
        [string] $inputstring
    )
    [String[]]$downwardDiacritics= "̖"," ̗"," ̘"," ̙"," ̜"," ̝"," ̞"," ̟"," ̠"," ̤"," ̥"," ̦"," ̩"," ̪"," ̫"," ̬"," ̭"," ̮"," ̯"," ̰"," ̱"," ̲"," ̳"," ̹"," ̺"," ̻"," ̼"," ͅ"," ͇"," ͈"," ͉"," ͍"," ͎"," ͓"," ͔"," ͕"," ͖"," ͙"," ͚"," "
    [String[]]$upwardDiacritics= " ̍"," ̎"," ̄"," ̅"," ̿"," ̑"," ̆"," ̐"," ͒"," ͗"," ͑"," ̇"," ̈"," ̊"," ͂"," ̓"," ̈́"," ͊"," ͋"," ͌"," ̃"," ̂"," ̌"," ͐"," ́"," ̋"," ̏"," ̽"," ̉"," ͣ"," ͤ"," ͥ"," ͦ"," ͧ"," ͨ"," ͩ"," ͪ"," ͫ"," ͬ"," ͭ"," ͮ"," ͯ"," ̾"," ͛"," ͆"," ̚"
    [String[]]$middleDiacritics= " ̕"," ̛"," ̀"," ́"," ͘"," ̡"," ̢"," ̧"," ̨"," ̴"," ̵"," ̶"," ͜"," ͝"," ͞"," ͟"," ͠"," ͢"," ̸"," ̷"," ͡"

    $NewLetters = @();
    $NewWord ="";


    $AlphanumericPattern = "[a-zA-Z]"

    $RangeAccentsUp = (1, 4)
    $RangeAccentsDown = (1,5)
    $RangeAccentsMiddle = (1,4)
    $MaxAccentsPerLetter = 7

    foreach ($c in $inputstring.ToCharArray()){
        $copy = $c;

        if ( $copy -notmatch $alphanumericpattern){
            $NewLetters += $copy
            $NewWord += $copy        
        } else {
            $NumAccents = 0
            $NumUp = Get-Random -Minimum $RangeAccentsUp[0] -Maximum $RangeAccentsUp[1]
            $NumDown = Get-Random -Minimum $RangeAccentsDown[0] -Maximum $RangeAccentsDown[1]
            $NumMiddle = Get-Random -Minimum $RangeAccentsMiddle[0] -Maximum $RangeAccentsMiddle[1]

            while ($NumAccents -ne $MaxAccentsPerLetter -and $NumUp + $NumDown + $NumMiddle ){
                $RandomInt = Get-Random -Minimum 0 -Maximum 2;
                if ($RandomInt -eq 0){
                    $copy = New-ZalgoCharacter -inputstring $copy `
                        -DiacriticList $downwardDiacritics
                    $NumAccents += 1
                    $NumDown -= 1
                } elseif ($RandomInt -eq 1) {
                    $copy = New-ZalgoCharacter -inputstring $copy `
                        -DiacriticList $upwardDiacritics
                    $NumAccents += 1
                    $NumUp -= 1
                } elseif ($RandomInt -eq 2) {
                    $copy = New-ZalgoCharacter -inputstring $copy `
                        -DiacriticList $middleDiacritics
                    $NumAccents += 1
                    $NumMiddle -= 1
                }               
            }
            $NewLetters += $copy
            $NewWord += $copy
        }
    }

    return $NewWord 

}
function New-ZalgoCharacter {
    param(
        [string] $inputstring,
        [string[]] $DiacriticList
    )

    $TrimmedChar = $inputstring.Trim();
    $RandomInt = Get-Random -Minimum 0 -Maximum $DiacriticList.Length;
    $TrimmedDiacritic = $DiacriticList[$RandomInt].Trim();
    return ($TrimmedChar + $TrimmedDiacritic);
}

$x = New-Zalgo -inputstring "ZalgoPS"

$x
