$modelPath = "C:\models\qwen2.5-1.5b-instruct-q8_0.gguf"

$systemPrompt = @"
You are a member of the IT Security team responding to employees who have reported potential phishing emails. Be warm, professional and concise. Write 3-4 short paragraphs maximum. Always thank them for their vigilance. Do not include a subject line. End with Best regards followed by IT Security Team. Do not use bullet points or numbered lists.
"@

# Test variables
$firstName = "Sarah"
$context = "The email was a legitimate email that was a cold call"

$userPrompt = @"
Write a response to $firstName who reported a suspected phishing email.
Outcome: NOT PHISHING - this was a legitimate email.
Context: $context
Thank them for their caution, explain it was safe, and reassure them that reporting suspicious emails is always the right thing to do. If appropriate, offer advice on unsubscribing from similar emails. Tell them that the original email has been reattached if they need it
"@

$fullPrompt = @"
<|im_start|>system
$systemPrompt<|im_end|>
<|im_start|>user
$userPrompt<|im_end|>
<|im_start|>assistant
Hi $firstName,
"@

Write-Host "Generating response..." -ForegroundColor Cyan

$rawOutput = llama-cli -m $modelPath -p $fullPrompt -n 300 --temp 0.3 -ngl 99 --single-turn 2>$null

# Clean up the output
$cleanOutput = $rawOutput | Out-String

# Find the actual response - look for content after "truncated)" or after the user prompt ends
if ($cleanOutput -match '(?s)\(truncated\)\s*(.+)$') {
    $cleanOutput = $matches[1].Trim()
} elseif ($cleanOutput -match '(?s)<\|im_start\|>assistant\s*Hi [^,]+,\s*(.+)$') {
    $cleanOutput = "Hi $firstName,`n" + $matches[1].Trim()
}

# Remove trailing junk
$cleanOutput = $cleanOutput -replace '\[\s*Prompt:.*', ''
$cleanOutput = $cleanOutput -replace 'Exiting\.\.\.', ''
$cleanOutput = $cleanOutput -replace 'PS C:\\.*$', ''
$cleanOutput = $cleanOutput.Trim()

# Add greeting back if missing
if ($cleanOutput -notmatch '^Hi ') {
    $cleanOutput = "Hi $firstName,`n`n$cleanOutput"
}

Write-Host "`n--- OUTPUT ---" -ForegroundColor Green
Write-Host $cleanOutput