#!/usr/bin/env bash
# Benchmark models on realistic Excel AI tasks
set -e

MODELS=("qwen2.5:1.5b" "qwen3.5:0.8b" "qwen3.5:2b" "phi4-mini" "llama3.2:3b")

PROMPTS=(
  "What is the capital of France? Reply with just the city name."
  "Classify the sentiment of this text as Positive, Negative, or Neutral. Reply with one word only.\n\nText: The product arrived late and was damaged, very disappointing."
  "Convert 72 degrees Fahrenheit to Celsius. Reply with just the number."
  "Translate to Spanish: Good morning, how are you?"
  "Summarize in one sentence: Apple reported Q4 revenue of 94.9 billion dollars, up 6 percent year over year, driven by strong iPhone and Services growth. Net income was 14.7 billion. CEO Tim Cook highlighted the installed base reaching an all-time high."
  "Extract the email address from this text. Reply with just the email.\n\nPlease contact our support team at help@acmecorp.com or call 555-0123."
  "What Excel formula would sum all values in column A? Reply with just the formula."
  "Is 17 a prime number? Reply Yes or No."
  "What is 15% of 230? Reply with just the number."
  "Categorize this expense: Uber ride to airport. Reply with one category name."
)

PROMPT_LABELS=(
  "Factual Q&A"
  "Sentiment"
  "Unit conversion"
  "Translation"
  "Summarization"
  "Extraction"
  "Excel formula"
  "Yes/No"
  "Math"
  "Categorize"
)

EXPECTED=(
  "Paris"
  "Negative"
  "22.2"
  "Buenos días"
  "*revenue*"
  "help@acmecorp.com"
  "=SUM(A:A)"
  "No"
  "34.5"
  "Transport"
)

echo "================================================================="
echo "  Excel AI Model Benchmark"
echo "================================================================="
echo ""

for model in "${MODELS[@]}"; do
  echo "--- $model ---"
  total_ms=0
  correct=0

  for i in "${!PROMPTS[@]}"; do
    prompt="${PROMPTS[$i]}"
    label="${PROMPT_LABELS[$i]}"

    start=$(python3 -c "import time; print(int(time.time()*1000))")

    response=$(curl -sk https://127.0.0.1:11435/v1/chat/completions \
      -H "Content-Type: application/json" \
      -d "$(printf '{"model":"%s","messages":[{"role":"system","content":"You are a helpful assistant embedded in Excel. Give concise answers suitable for spreadsheet cells."},{"role":"user","content":"%s"}],"temperature":0.1,"max_tokens":80}' "$model" "$prompt")" 2>/dev/null)

    end=$(python3 -c "import time; print(int(time.time()*1000))")
    ms=$((end - start))
    total_ms=$((total_ms + ms))

    answer=$(echo "$response" | python3 -c "import sys,json; print(json.load(sys.stdin)['choices'][0]['message']['content'].strip())" 2>/dev/null || echo "ERROR")

    # Truncate long answers for display
    display=$(echo "$answer" | head -1 | cut -c1-60)

    printf "  %-16s %5dms  %s\n" "$label" "$ms" "$display"
  done

  avg_ms=$((total_ms / ${#PROMPTS[@]}))
  printf "  %-16s %5dms avg (%dms total)\n" "AVERAGE" "$avg_ms" "$total_ms"
  echo ""
done
