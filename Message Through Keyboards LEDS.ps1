# Parameters
Param(
    [Parameter(Mandatory = $True, HelpMessage = "Enter the sentence to be written with LEDs")]
    [string]$Sentence,
    [Parameter(Mandatory = $False, HelpMessage = "How much time the LEDs will blink before sentence")]
    [int]$BlinkTime = 1,
    [Parameter(Mandatory = $False, HelpMessage = "Enable Vertbosity to see what is the octal value for each letter")]
    [switch]$VerboseOutput
)

# Reseting all keyboard LEDs
function Reset-LEDS {
    if ([System.Windows.Forms.Control]::IsKeyLocked('NumLock'))
    { $keyboard.SendKeys("{NUMLOCK}") }
    if ([System.Windows.Forms.Control]::IsKeyLocked('CapsLock'))
    { $keyboard.SendKeys("{CAPSLOCK}") }
    if ([System.Windows.Forms.Control]::IsKeyLocked('Scroll'))
    { $keyboard.SendKeys("{SCROLLLOCK}") }
}

# Getting Users Attention
function Blink-LEDS {
    Start-Sleep -Seconds 1
    for ($i = 0; $i -lt ($BlinkTime * 6) ; $i++) {
        $keyboard.SendKeys("{NUMLOCK}")
        $keyboard.SendKeys("{CAPSLOCK}")
        $keyboard.SendKeys("{SCROLLLOCK}")
        Start-Sleep -Seconds 0.55
    }
    # Safety mechanism
    Reset-LEDS
    Start-Sleep -Seconds 0.5
}

# End of letter blinking
function EndLetter-LEDS{
    Reset-LEDS
    Start-Sleep -Seconds 0.55
    for ($i = 0; $i -lt 4; $i++){
        $keyboard.SendKeys("{NUMLOCK}"); Start-Sleep -Seconds 0.55
        $keyboard.SendKeys("{CAPSLOCK}"); Start-Sleep -Seconds 0.55
        $keyboard.SendKeys("{SCROLLLOCK}"); Start-Sleep -Seconds 0.55
    }
    Reset-LEDS
}

# --------------------------------------------------------

# Creating the keyboard object
$keyboard = New-Object -ComObject WScript.Shell
# Adding System.Windows.Forms Assembly
Add-Type -AssemblyName System.Windows.Forms

# Saving current keyboard state
$NUM_STATE = [System.Windows.Forms.Control]::IsKeyLocked('NumLock')
$CAPS_STATE = [System.Windows.Forms.Control]::IsKeyLocked('CapsLock')
$SCROLL_STATE = [System.Windows.Forms.Control]::IsKeyLocked('Scroll')

# Reseting all keyboard LEDs
# Getting Users Attention and starting sentence
Reset-LEDS
Blink-LEDS

# Sending encrypted sentence via the keyboard LEDs
for ($i = 0; $i -lt $Sentence.Length; $i++)
{
    EndLetter-LEDS
    # Converting key to octal number
    $Key = [byte][char]$Sentence[$i]
    $Key = [Convert]::ToString($Key,8)
    $KeyOctal = @()
    $KeyOctal += [Math]::Floor($Key/100)
    $KeyOctal += [Math]::Floor($Key/10)%10
    $KeyOctal += $Key%10
    if ($VerboseOutput.IsPresent)
    {Write-Host $KeyOctal}

    # Looping through all number in key
    foreach ($Number in $KeyOctal)
    {
        $KeyOctalInBinary = [Convert]::ToString($Number,2)

        $IsZero = $True
        if (([Math]::Floor($KeyOctalInBinary/100)) -eq 1)
        { $keyboard.SendKeys("{NUMLOCK}"); $NotZero = $False }
        if (([Math]::Floor($KeyOctalInBinary/10)%10) -eq 1)
        { $keyboard.SendKeys("{CAPSLOCK}"); $NotZero = $False }
        if (($KeyOctalInBinary%10) -eq 1)
        { $keyboard.SendKeys("{SCROLLLOCK}"); $NotZero = $False }

        if ($IsZero -eq $True) {Start-Sleep -Seconds 1}
        Start-Sleep -Seconds 0.55
        Reset-LEDS
        Start-Sleep -Seconds 0.55
    }
}
EndLetter-LEDS

# Reseting all keyboard LEDs
# Getting Users Attention and ending sentence
Reset-LEDS
Blink-LEDS

# Returning original keyboard state
if ($NUM_STATE)
{ $keyboard.SendKeys("{NUMLOCK}") }
if ($CAPS_STATE)
{ $keyboard.SendKeys("{CAPSLOCK}") }
if ($SCROLL_STATE)
{ $keyboard.SendKeys("{SCROLLLOCK}") }
