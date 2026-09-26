from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# 1. FUNCTIONAL TEST CASES
# ============================================================

functional_test_cases = [
    [
        "TC-001",
        "Web Extension Open/Launch",
        "Chrome installed; EchoGPT extension installed",
        "Open Chrome and launch EchoGPT extension",
        "Extension opens successfully",
        "Extension opened successfully",
        "PASS"
    ],
    [
        "TC-002",
        "Web New Chat",
        "Extension is open",
        "Click New Chat",
        "A fresh chat should open",
        "Fresh chat opened",
        "PASS"
    ],
    [
        "TC-003",
        "Web Chat Message Send",
        "Chat screen is open",
        "Enter a message and send it",
        "Message should be sent and an AI response should appear",
        "Message sent and response appeared",
        "PASS"
    ],
    [
        "TC-004",
        "Web Model List Open",
        "Chat screen is open",
        "Open the model selector",
        "Available AI models should be displayed",
        "Model list displayed",
        "PASS"
    ],
    [
        "TC-005",
        "Web Model Selection",
        "Model list is open",
        "Select an available model",
        "Selected model should be applied",
        "Selected model applied",
        "PASS"
    ],
    [
        "TC-006",
        "Web Model Switching",
        "Chat screen is open",
        "Switch between available models",
        "Selected model should change successfully",
        "Model switched successfully",
        "PASS"
    ],
    [
        "TC-007",
        "Web Chat Input Send State",
        "Chat screen is open",
        "Type text and then clear the text",
        "Send state should activate with text and deactivate when empty",
        "Send state changed correctly",
        "PASS"
    ],
    [
        "TC-008",
        "Web Multiline Input",
        "Chat screen is open",
        "Enter multiline text using Shift+Enter",
        "A new line should be inserted without sending the message",
        "Multiline input worked",
        "PASS"
    ],
    [
        "TC-009",
        "Web File Attachment Picker",
        "Chat screen is open",
        "Click attachment icon, open file picker, then cancel",
        "File picker should open and Cancel should return to chat",
        "File picker opened and Cancel returned to chat",
        "PASS"
    ],
    [
        "TC-010",
        "Web History Open",
        "Extension is open",
        "Open History",
        "Previous conversations should be displayed",
        "History displayed",
        "PASS"
    ],
    [
        "TC-011",
        "Web Reopen Previous Chat",
        "History contains at least one conversation",
        "Open a previous conversation",
        "Previous conversation should open with its content",
        "Previous conversation reopened",
        "PASS"
    ],
    [
        "TC-012",
        "Web Translate Text",
        "Translate feature is open",
        "Enter text and translate it",
        "Translation result should be displayed",
        "Translation result displayed",
        "PASS"
    ],
    [
        "TC-013",
        "Web Translate Result Copy",
        "Translation result is available",
        "Click Copy and paste the result elsewhere",
        "Copied text should exactly match the displayed translation",
        "Copied result matched exactly",
        "PASS"
    ],
    [
        "TC-014",
        "Web Read Webpage",
        "Read feature is open",
        "Provide a valid webpage link and read it",
        "Webpage should be processed and a result should be displayed",
        "Webpage was processed and result displayed",
        "PASS"
    ],
    [
        "TC-015",
        "Web Write Validation",
        "Write feature is open",
        "Leave the topic field empty",
        "Generate should remain disabled or validation should be displayed",
        "Validation/disabled state worked",
        "PASS"
    ],
    [
        "TC-016",
        "Web Write Options",
        "Write feature is open",
        "Change format, tone, length and language options",
        "Available options should be selectable",
        "Options changed successfully",
        "PASS"
    ],
    [
        "TC-017",
        "Web Compare Model Selection",
        "A chat response is available",
        "Open Compare and select a model",
        "Selected model should be added for comparison",
        "Model selection worked",
        "PASS"
    ],
    [
        "TC-018",
        "Web MCP Connector Validation",
        "MCP feature is open",
        "Test empty, invalid and valid URL inputs",
        "Appropriate validation/error messages should appear",
        "Validation messages appeared",
        "PASS"
    ],
    [
        "TC-019",
        "Web Pin / Unpin Extension",
        "Extension is installed",
        "Pin and then unpin the extension",
        "Pin state should change successfully",
        "Pin/unpin worked",
        "PASS"
    ],
    [
        "TC-020",
        "Web Theme / Light Mode",
        "Extension is open",
        "Change the available extension theme setting",
        "Theme should change successfully",
        "Theme setting worked",
        "PASS"
    ],

    # ---------------- ANDROID ----------------

    [
        "TC-021",
        "Android New Chat",
        "Android application is open",
        "Tap New Chat",
        "A fresh chat should open",
        "Fresh chat opened",
        "PASS"
    ],
    [
        "TC-022",
        "Android History",
        "Android application is open",
        "Open History",
        "Previous conversations should be displayed",
        "History displayed",
        "PASS"
    ],
    [
        "TC-023",
        "Android Reopen Previous Conversation",
        "History contains a conversation",
        "Open a previous conversation",
        "Previous conversation should reopen",
        "Previous conversation reopened",
        "PASS"
    ],
    [
        "TC-024",
        "Android Subscription Page",
        "Android application is open",
        "Open Subscription",
        "Subscription page should open",
        "Subscription page opened",
        "PASS"
    ],
    [
        "TC-025",
        "Android Subscription Duration Selection",
        "Subscription page is open",
        "Select 1M, 3M, 6M and 1Y options",
        "Selected duration should change",
        "Duration selection worked",
        "PASS"
    ],
    [
        "TC-026",
        "Android Image Studio Open",
        "Android application is open",
        "Open Image Studio",
        "Image Studio should open",
        "Image Studio opened",
        "PASS"
    ],
    [
        "TC-027",
        "Android Image Studio Options",
        "Image Studio is open",
        "Change available Image Studio options",
        "Options should be selectable",
        "Options changed successfully",
        "PASS"
    ],
    [
        "TC-028",
        "Android Compare Module",
        "Android application is open",
        "Open Compare and inspect model selection",
        "Compare module and model selection should open",
        "Compare module opened",
        "PASS"
    ],
    [
        "TC-029",
        "Android Connectors Module",
        "Android application is open",
        "Open Connectors",
        "Connectors page should open",
        "Connectors page opened",
        "PASS"
    ],
    [
        "TC-030",
        "Android Support Page",
        "Android application is open",
        "Open Support",
        "Support page should open",
        "Support page opened",
        "PASS"
    ],
    [
        "TC-031",
        "Android Share EchoChat",
        "Android application is open",
        "Tap Share EchoChat",
        "Native share sheet should open",
        "Share sheet opened",
        "PASS"
    ],
    [
        "TC-032",
        "Android Discord Navigation",
        "Android application is open",
        "Open Discord and return to the application",
        "Discord navigation should work and Back should return",
        "Navigation worked",
        "PASS"
    ],
    [
        "TC-033",
        "Android Model Selection",
        "Chat screen is open",
        "Open model selector and choose a model",
        "Selected model should be applied",
        "Model selected",
        "PASS"
    ],
    [
        "TC-034",
        "Android Model Switching",
        "Chat screen is open",
        "Switch between available models",
        "Selected model should change",
        "Model switched",
        "PASS"
    ],
    [
        "TC-035",
        "Android Chat Input State",
        "Chat screen is open",
        "Type and clear a message",
        "Input/send state should update correctly",
        "Input state worked",
        "PASS"
    ],
    [
        "TC-036",
        "Android Long Text Input",
        "Chat screen is open",
        "Paste a long text",
        "Long text should remain usable without UI failure",
        "Long text input worked",
        "PASS"
    ],
    [
        "TC-037",
        "Android Multiline Input",
        "Chat screen is open",
        "Enter multiline text",
        "Multiple lines should be accepted",
        "Multiline input worked",
        "PASS"
    ],
    [
        "TC-038",
        "Android Keyboard & Back Behavior",
        "Chat input is active",
        "Open keyboard and press Back",
        "Keyboard should close and screen should remain usable",
        "Back behavior worked",
        "PASS"
    ],
    [
        "TC-039",
        "Android Screen Rotation",
        "Android application is open",
        "Rotate the device/simulator",
        "Application should remain usable and state should be retained",
        "Rotation worked",
        "PASS"
    ],
    [
        "TC-040",
        "Android App Background/Resume & Restart",
        "Android application is open",
        "Background the app, resume it, then restart it",
        "App should return/restart without unexpected failure",
        "Resume and restart worked",
        "PASS"
    ],
]


# ============================================================
# 2. BUG REPORTS
# ============================================================

bug_reports = [
    [
        "WEB-BUG-001",
        "Translate language swap (⇄) button does not work",
        "The language swap control in the Translate feature does not exchange the source and target languages.",
        """1. Open Translate.
2. Set source and target languages.
3. Click the ⇄ swap button.""",
        "Source and target languages should exchange places.",
        "No visible change occurs and the swap control does not work.",
        "Medium",
        "High",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],

    [
        "WEB-BUG-002",
        "Read Aloud produces no audio",
        "Read Aloud can be triggered, but no audio playback is produced.",
        """1. Open a chat response.
2. Click Read Aloud.
3. Wait 10–15 seconds.""",
        "Audio playback should start.",
        "No audio playback or visible playback indicator occurs.",
        "Medium",
        "High",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],

    [
        "WEB-BUG-003",
        "Settings intermittently redirects to Sign In",
        "Opening Settings intermittently redirects to the Sign In screen instead of opening the Settings page.",
        """1. Open the EchoGPT extension.
2. Open Settings.
3. Repeat the action during separate attempts.""",
        "Settings should open directly for the authenticated user.",
        "Settings redirected to Sign In on two earlier attempts, although it later opened correctly.",
        "Medium",
        "High",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Intermittent"
    ],

    [
        "WEB-BUG-004",
        "Compare with selected model does not generate response",
        "After selecting a comparison model, the selected model does not generate a comparison response.",
        """1. Send: What is 2 + 2?
2. Open Compare with.
3. Select DeepSeek V4 Pro.
4. Wait 10–15 seconds.""",
        "The selected model should generate a comparison response.",
        "The model selection occurs, but no comparison response is generated.",
        "Medium",
        "High",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],

    [
        "WEB-BUG-005",
        "New Chat does not clear previously attached file",
        "An attachment from the previous chat remains visible after starting a New Chat.",
        """1. Attach BUG-003.png.
2. Click New Chat.
3. Inspect the chat input area.""",
        "New Chat should start with no attachment from the previous chat.",
        "The previous attachment remains visible.",
        "Medium",
        "High",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],

    [
        "WEB-BUG-006",
        "New Chat retains unsent message draft from previous chat",
        "An unsent chat draft remains after starting a New Chat and can also persist while navigating through History.",
        """1. Type an unsent message.
2. Click New Chat.
3. Inspect the chat input.""",
        "New Chat should provide an empty message input.",
        "The previous unsent draft remains.",
        "Medium",
        "Medium",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],

    [
        "WEB-BUG-007",
        "Previous translation result remains after clearing source input",
        "The previous Translate result remains visible after the source text is completely removed.",
        """1. Translate "Hello".
2. Wait for the translation result.
3. Clear the source text completely.""",
        "The translation result should clear or reset when the source input is removed.",
        "The previous translation result remains visible.",
        "Medium",
        "Medium",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],

    [
        "WEB-BUG-008",
        "New Chat retains previous translation result",
        "The Translate result from a previous chat remains after starting a New Chat.",
        """1. Generate a translation.
2. Click New Chat.
3. Open Translate again.""",
        "Translate should start with a fresh state.",
        "The previous translation result is still displayed.",
        "Medium",
        "Medium",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],

    [
        "WEB-BUG-009",
        "New Chat retains previous Read result",
        "The previous Read result remains after starting a New Chat and returning to the Read feature.",
        """1. Read a valid TXT file.
2. Click New Chat or open another feature.
3. Return to Read.""",
        "Read should start with a fresh result state.",
        "The previous Read result text remains. The old file itself is not attached.",
        "Medium",
        "Medium",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],

    [
        "WEB-BUG-010",
        "New Chat retains unsent Write input",
        "An unsent Write topic remains populated after starting a New Chat.",
        """1. Open Write.
2. Enter "Testing Write State".
3. Do not generate.
4. Click New Chat.
5. Open Write again.""",
        "Write input should be reset for the new chat.",
        'The text "Testing Write State" remains in the Write input.',
        "Medium",
        "Medium",
        "Windows 11; Google Chrome; EchoGPT Chrome Extension",
        "Confirmed"
    ],
]


# ============================================================
# 3. CREATE EXCEL WORKBOOK
# ============================================================

file_name = "EchoGPT_SQA_Assignment_Functional_and_Bug_Reports.xlsx"

workbook = Workbook()


# ============================================================
# 4. FUNCTIONAL TEST CASE SHEET
# ============================================================

ws_functional = workbook.active
ws_functional.title = "Functional Test Cases"

functional_headers = [
    "Test Case ID",
    "Feature",
    "Preconditions",
    "Test Steps",
    "Expected Result",
    "Actual Result",
    "Status"
]

ws_functional.append(functional_headers)

for test_case in functional_test_cases:
    ws_functional.append(test_case)


# ============================================================
# 5. BUG REPORT SHEET
# ============================================================

ws_bug = workbook.create_sheet("Bug Reports")

bug_headers = [
    "Bug ID",
    "Title",
    "Description",
    "Steps to Reproduce",
    "Expected Result",
    "Actual Result",
    "Severity",
    "Priority",
    "Environment",
    "Status"
]

ws_bug.append(bug_headers)

for bug in bug_reports:
    ws_bug.append(bug)


# ============================================================
# 6. STYLING
# ============================================================

header_fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

header_font = Font(
    bold=True,
    color="FFFFFF",
    size=11
)

thin_border = Border(
    left=Side(style="thin", color="B7B7B7"),
    right=Side(style="thin", color="B7B7B7"),
    top=Side(style="thin", color="B7B7B7"),
    bottom=Side(style="thin", color="B7B7B7")
)


def format_sheet(ws, column_widths):

    # Freeze first row
    ws.freeze_panes = "A2"

    # Add filter
    ws.auto_filter.ref = ws.dimensions

    # Header styling
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = thin_border

    # Body styling
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )
            cell.border = thin_border

    # Column widths
    for index, width in enumerate(column_widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width

    # Header height
    ws.row_dimensions[1].height = 35

    # Body row height
    for row_number in range(2, ws.max_row + 1):
        ws.row_dimensions[row_number].height = 70


# Functional sheet widths
functional_widths = [
    16, 32, 38, 55, 48, 48, 14
]

format_sheet(
    ws_functional,
    functional_widths
)


# Bug sheet widths
bug_widths = [
    17, 42, 55, 60, 50,
    58, 14, 14, 50, 16
]

format_sheet(
    ws_bug,
    bug_widths
)


# ============================================================
# 7. ADD SUMMARY INFORMATION
# ============================================================

ws_functional.sheet_view.showGridLines = False
ws_bug.sheet_view.showGridLines = False


# ============================================================
# 8. SAVE FILE
# ============================================================

workbook.save(file_name)

print("=" * 70)
print("Excel file generated successfully!")
print("=" * 70)
print(f"File: {file_name}")
print()
print("Sheets:")
print("1. Functional Test Cases - 40 test cases")
print("2. Bug Reports - 10 genuine issues")
print()
print("The workbook is ready for the assignment.")