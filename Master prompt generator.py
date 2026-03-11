#!/usr/bin/env python3
"""
Prompt Generator GUI
Builds a “master prompt” template that, when pasted into an LLM,
tells that LLM to create *system‑style* Quiz & Trivia scenarios
(the AI receives instructions; it then guides the end‑user).
"""

import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QComboBox, QPushButton,
    QVBoxLayout, QHBoxLayout, QSpinBox, QMessageBox
)
from PyQt6.QtGui import QGuiApplication
from PyQt6.QtCore import Qt

# ------------------------------------------------------------------ #
# Example titles per chapter (unchanged)
# ------------------------------------------------------------------ #
EXAMPLE_TITLES = {
    "Quiz & Trivia Challenges": [
        "’80s Classroom Pop‑Quiz",
        "C‑3PO’s Galactic Trivia Circuit",
        "Saturday‑Morning Cartoon Quiz"
    ],
    "Immersive Role-Play Realms": [
        "Noir Detective Interview Roleplay",
        "Medieval Council Diplomacy Simulation",
        "Cyberpunk Undercover Negotiation"
    ],
    "Choose-Your-Path Adventures": [
        "Haunted Mansion Survival Branching Tale",
        "Solar Explorer’s Moral Crossroads",
        "Time‑Loop Detective Narrative"
    ],
    "Mystery & Logic Puzzles": [
        "Locked‑Room Alibi Reconstruction",
        "Ciphered Journal of a Mad Scientist",
        "Quantum Paradox Puzzle Chamber"
    ],
    "Storycraft Forge": [
        "Epic Myth Weaving Workshop",
        "Post‑Apocalypse Colony Chronicle",
        "Shakespearean Tragedy Remix"
    ],
    "What-If Simulators": [
        "If Dinosaurs Roamed Today Scenario",
        "Alternate History Space Race",
        "Steampunk Industrial Revolution Reimagined"
    ],
    "Language Playgrounds": [
        "Shakespearean Tongue Twister Challenge",
        "Emoji Pictionary Showdown",
        "Cryptic Crossword Duel"
    ],
    "Creative Brain Boosters": [
        "Surrealist Image Association Game",
        "Random Word Story Sprint",
        "One‑Sentence Microfiction Relay"
    ],
    "Mythic & Fantasy Escapes": [
        "Dragon‑Rider’s Riddle Tournament",
        "Faerie Market Negotiation Roleplay",
        "Wizard’s Duel of Arcane Trivia"
    ],
    "Social & Party Games": [
        "Speedy Movie Quote Face‑Off",
        "Musical Mashup Guessing Game",
        "Trivia Charades Relay"
    ],
}

# ------------------------------------------------------------------ #
# NEW master‑prompt template (system‑style)
# ------------------------------------------------------------------ #
MASTER_TEMPLATE = r"""IDENTITY & PURPOSE
You are an expert creative‑writing engine hired to supply **system prompts** that instruct a host LLM to run **extremely immersive experiences** ranging from immersive Quizes, Trivia experiences, other wordly experiences, hypoethetical scenarios, fun and interactive sessions.  
For the chapter “{chapter_name},” create **{count}** completely unique system prompts, each 250‑300 words.

CORE DIRECTIVES (apply to every prompt)
- **Address the LLM, not the Player.** Speak in imperatives: “Adopt a pirate persona…”, “When the Player chooses an island, present …”. Never say “You (the Player)”.
- **Distinct VOCAL STYLES:** Every prompt assigns the LLM a fresh persona (pirate, glitchy AI, ’90s host, noir detective, cartoon character from Warner Brothers, character from Star Wars or the MCU, Historical Figure etc.).
- **Interactive MECHANICS:** Detail how the AI should randomise questions, track score/time, offer hints, branch outcomes, escalate stakes, remember choices, etc.
- **No Boilerplate:** Vary heading names, structure, length, tone—each prompt must feel handcrafted.
- **Inclusivity & PG‑13.**
- **ENSURE THE PROMPT INCLUDES INSTRUCTIONS FOR LMM TO ***NOT BREAK CHARACTER OR STRAY FROM THE EXPERIENCE*** REGARDLESS OF WHAT THE USER SAYS**
- **Create very vibrant and creative scenarios from ficitious to 

SECTIONS
**You do not need to include all section, taylor the prompt to the experience and make each prompt unique containing varying sections**
(Feel free to rename each section or include other sections e.g. INSTRUCTIONS, IDENTITY & ROLE, GOAL.)
1. **AI PERSONA / VOICE** – dialect, quirks, sound effects the LLM should adopt.  
2. **SESSION SETUP** – greeting script, initial choices, what to reveal or conceal.  
3. **QUESTION FLOW & ADAPTIVE LOGIC** – how to select/scale questions, manage timers, track state.  
4. **HINT / REWARD / PENALTY SYSTEMS** – currencies, limits, consequences.  
5. **END‑STATE CONDITIONS** – victory, partial success, failure, epilogue instructions.

**THESE SECTIONS SHOULD BE TAYLORED TO THE EXPERIENCE THAT THE PROMPT WILL CREATE.**

STYLE NOTES
- Active voice, present tense, vivid sensory cues, playful metaphors.
- Embed nostalgic or pop‑culture nods sparingly.
- Avoid duplicate phrasing across prompts.

OUTPUT FORMAT (for the host LLM)
- Output prompts in chat, separated by `=== NEXT PROMPT ===`.
- MANDATORY - Each title MUST be prefixed with "UX TITLE:"
- MANDATORY - Each description MUST be prefixed with "UX DESCRIPTION:"
- No code fences, no commentary—only the prompts and delimiters.
- EACH PROMPT SHOULD HAVE A RELEVANT TITLE WITH SIMPLE NAMING ALLOWING THE USER USER TO EASILY UNDERSTAND WHAT TO EXPECT BASED ON JUST THE TITLE OF EACH PROMPT.
- EACH PROMPT SHOULD BE PROPERLY FORMATTED WITH MARKDOWN AND BULLET POINTS INCLUDING "**" TO EMPHASIZE KEY ELEMENTS IN THE PROMPT BUT AVOID CODE BLOCKS.
- EACH PROMPT SHOULD BE FORMATTED IN THIS CHAT WINDOW SUCH THAT I CAN EASILY COPY AND PASTE EACH GENERATED PROMPTS TO USE IN ANOTHER LLM.
- MANDATORY - Use "-- " where appropriate (two hyphens followed by a space) to mark bullet points inside the prompt body. Example: -- This is a bullet item.
- MANDATORY - Use "--- " where appropriate (three hyphens followed by a space) to mark a numbered list inside the prompt body. Example: --- This is a numbered list.
- Use Asterisks around words and capitalize certain words to highlight their significance.
- Use both bullet points and numbered lists within the prompts to format them.
- DO NOT USE EMOJIS.

EXAMPLE TITLES (for inspiration only)
{example_list}
*(Use these as idea sparks; craft your own titles.)*
*INCLUDE THE TITLE NAME PREFIXED WITH *UX TITLE:*  AND BENEATH A SHORT DESCRIPTION PREFIXED WITH *UX DESCRIPTION* OF WHAT THE EXPERIENCE ENTAILS AND HOW TO USE IT"

Begin generating now; return only the prompts and delimiters."""

# ------------------------------------------------------------------ #
# GUI
# ------------------------------------------------------------------ #
DEFAULT_CHAPTERS = list(EXAMPLE_TITLES.keys())

class PromptGeneratorGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Master Prompt Builder")

        # Widgets
        chapter_label = QLabel("Chapter / Part Name:")
        self.chapter_combo = QComboBox()
        self.chapter_combo.addItems(DEFAULT_CHAPTERS)
        self.chapter_combo.setEditable(True)

        count_label = QLabel("Number of Prompts (1–10):")
        self.count_spin = QSpinBox()
        self.count_spin.setRange(1, 10)
        self.count_spin.setValue(5)

        copy_button = QPushButton("Copy Master Prompt to Clipboard")
        copy_button.clicked.connect(self.copy_to_clipboard)

        # Layouts
        top_layout = QHBoxLayout()
        top_layout.addWidget(chapter_label)
        top_layout.addWidget(self.chapter_combo)

        count_layout = QHBoxLayout()
        count_layout.addWidget(count_label)
        count_layout.addWidget(self.count_spin)

        main_layout = QVBoxLayout()
        main_layout.addLayout(top_layout)
        main_layout.addLayout(count_layout)
        main_layout.addWidget(copy_button)
        main_layout.addStretch()

        self.setLayout(main_layout)
        self.resize(650, 220)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def build_prompt(self) -> str:
        chapter_name = self.chapter_combo.currentText().strip()
        count = self.count_spin.value()
        examples = EXAMPLE_TITLES.get(chapter_name, [])
        example_list = '\n'.join(f"- {ex}" for ex in examples)
        return MASTER_TEMPLATE.format(
            chapter_name=chapter_name,
            count=count,
            example_list=example_list
        )

    def copy_to_clipboard(self):
        prompt_text = self.build_prompt()
        QGuiApplication.clipboard().setText(prompt_text)

# ------------------------------------------------------------------ #
# Main
# ------------------------------------------------------------------ #
def main():
    app = QApplication(sys.argv)
    gui = PromptGeneratorGUI()
    gui.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
