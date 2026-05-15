# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a C# Windows Forms AOI (Automated Optical Inspection) application for industrial bushing manufacturing quality control. The system integrates Basler cameras, PLC communication, YOLO object detection, and anomaly detection using TensorRT to inspect products in real-time.

**Target Framework:** .NET Framework 4.8 (x64)
**Main Entry:** `Program.cs` → `Form1`
**Build Output:** `bin\x64\Release\peilin.exe`

## Build and Run

### Building the Project
```bash
# Build in Visual Studio
msbuild peilin.sln /p:Configuration=Release /p:Platform=x64

# Or use Visual Studio IDE (recommended)
# Set configuration to "Release" and platform to "x64"
```

### Running the Application
- Execute `bin\x64\Release\peilin.exe`
- The application requires a SQLite database at `.\setting\mydb.sqlite`
- Smart card dongle authentication is enabled by default (can be disabled via `app.usekey`)

### External Dependencies
- **Python Server for YOLO**: Start via batch files in `bin\x64\Release\yolo_*.bat`
- **TensorRT DLLs**: `AD_TRT_dll1.dll` through `AD_TRT_dll4.dll` for anomaly detection
- **Basler Pylon SDK**: Required for camera operations

## System Architecture

### Multi-Station Design
The system supports up to 4 inspection stations simultaneously:
- Each station has its own Basler camera (configured via `Camera0.cs`)
- Each station runs independent AI inference (TensorRT + YOLO)
- Shared PLC communication coordinates all stations via serial port (Modbus protocol)

### Core Processing Pipeline
```
Camera Trigger → Image Capture → Position Detection →
ROI Extraction → AI Inference (Anomaly + YOLO) →
Defect Classification → Result Storage → PLC Signal
```

### Key Architectural Components

**Form1.cs** (17,000+ lines)
- Main application form and system orchestration
- Contains the `app` static class with global state (queues, counters, parameters)
- Image processing pipeline coordination
- PLC communication handlers

**Camera0.cs**
- Basler Pylon camera wrapper
- Multi-camera management (`m_BaslerCameras` array)
- Event-driven image grabbing with monitoring/logging

**anomalyTensorRT.cs**
- Wrapper for 4 TensorRT DLL instances (`AD_TRT_dll1` through `AD_TRT_dll4`)
- Anomaly detection model loading and inference
- Returns anomaly heatmap (Mat) and score (float)

**YoloDetection.cs**
- HTTP client for Python-based YOLO server
- Model loading, warmup, and batch detection
- Communicates with localhost servers on ports 5001-5004

**PLC_Test.cs**
- Modbus serial communication with PLC
- Handles trigger signals, counters (D803, D805, D807), and air blow commands
- Priority queue system for PLC commands to prevent congestion

**ParameterConfigForm.cs / ParameterSetupManager.cs**
- Parameter management UI with 3-zone system:
  - Reference Zone (source part number params, read-only)
  - Added-Unmodified Zone (copied but not yet edited)
  - Added-Modified Zone (edited parameters ready for save)
- Calibration tool integration for position/circle/pixel parameters

**algorithm.cs**
- SafeROI helper methods for boundary-safe image region extraction

### Database Schema (SQLite)
Located at `.\setting\mydb.sqlite` with the following tables:
- **Cameras**: Camera parameters (exposure, gain, delay) per part number
- **params**: Algorithm parameters per part number
- **Types**: Part number information
- **DefectTypes**: Global defect type definitions
- **DefectChecks**: Part-number-specific defect detection mappings
- **DefectCounts**: Defect statistics for reporting
- **Totals**: Production counters
- **Blows**: Air blow sorting parameters
- **Users**: User authentication

## Code Style Requirements

### Communication Language
**All AI assistant responses must be in Traditional Chinese (繁體中文) only.**
- Every response, explanation, and communication with users must use Traditional Chinese
- Do not use Simplified Chinese (简体中文) or English in responses
- Technical terms may remain in English when no standard Chinese translation exists

### No Emoji Usage
**Do not use emoji in any responses.**
- All communication must be text-only without emoji characters
- Use clear textual descriptions instead of emoji for emphasis or emotion

### Language & Comments
- **All comments must be in Traditional Chinese (繁體中文)**

### Database Operations (LinqToDB)
Always use "update-first, insert-if-none" pattern to avoid primary key conflicts:
```csharp
// 嘗試更新現有記錄
int updatedRows = db.table.Where(條件).Set(欄位, 值).Update();

// 若無更新記錄則插入新記錄
if (updatedRows == 0)
{
    db.table.Value(欄位, 值).Insert();
}
```
- Avoid creating entity objects for Insert; use `.Value()` syntax
- Always use `using` statements for database connections
- Include proper error handling for all DB operations

### Memory Management
**Critical:** This application processes high-resolution images at industrial speeds. Poor memory management causes crashes.

- **OpenCV Mat objects**: Must be disposed immediately after use
- **System.Drawing.Bitmap**: Must call `.Dispose()` explicitly
- Use `using` statements for image processing operations
- Long-lived large objects should use compressed storage
- The application uses `_disposeQueue` and `_disposeSemaphore` for async disposal

### Namespace Disambiguation
**Always fully qualify `System.Drawing.*` types even when `using System.Drawing` is present:**
```csharp
// CORRECT
System.Drawing.Point point = new System.Drawing.Point(x, y);
System.Drawing.Bitmap bitmap = ...;

// WRONG (conflicts with OpenCvSharp)
Point point = new Point(x, y);
```
This is because OpenCvSharp also defines `Point`, `Size`, `Rect`, etc.

## Important Global State (`app` class)

The `app` static class in Form1.cs contains critical shared state:
- **Queues**: `Queue_Bitmap1-4`, `Queue_Save`, `Queue_Send`, `Queue_Show` for pipeline stages
- **Counters**: `app.counter`, `app.dc` for production tracking
- **Parameters**: `app.param`, `app.models`, `app.metas`, `app.pos` loaded from database
- **Wait Handles**: `_wh1-4`, `_sv`, `_reader`, `_AI`, `_show` for thread synchronization
- **System State**: `app.currentState` enum (tracks system operational state)
- **PLC Queues**: `pendingOK1`, `pendingOK2`, `pendingPushOK1`, `pendingPushOK2` for push tracking

## PLC Communication Notes

### Counting Mechanism
- Push rod counts are managed by PLC (registers D803, D805, D807)
- Software displays values read from PLC
- `app.counter["stop" + camID]` tracks software-side counts
- `SAMPLE_ID` is derived from `app.counter`
- Known issue: Emergency stop can cause +1-2 count mismatch between software report and physical count

### Command Priority
- Air blow commands have highest priority to prevent queue congestion
- Recent commits improved PLC SerialPort receive handling and removed excessive `GC.Collect()` calls

## Recent Major Changes

### ParameterConfigForm Refactoring
See `IMPLEMENTATION_SUMMARY.md` for detailed documentation:
- New 3-zone UI architecture (Reference → Added-Unmodified → Added-Modified)
- Fixed `objBias` parameter miscategorization (moved to Detection category)
- Fixed position parameters disappearing after calibration
- Visual status indicators via background colors
- Calibration tools (Circle, Contrast, Pixel, White, ObjectBias, GapThresh) now preserve reference zone

### Memory and Performance
- Removed excessive `GC.Collect()` calls (see commit: "避免PLC指令堆積/修復記憶體錯誤")
- Improved PLC SerialPort event handling
- Set offline mode to false by default

## Session and Configuration Files

### Parameter Session Files
- Location: `bin\x64\Release\setting\param_session_*.json`
- Format: JSON with part number as key
- Contains all parameter categories for quick session restore
- Example: `param_session_10201120714TP.json`

### Model File Structure
- Anomaly models: `bin\x64\Release\models\{PartNumber}_{in|out|1|2}.pt`
- YOLO models: Python .pt files served via HTTP
- TensorRT engines: Previously stored, now using .pt models with dynamic loading
- Metadata: `.json` files accompanying each model for normalization parameters

## Testing and Calibration

### Calibration Tools (accessible from ParameterConfigForm)
- **CircleCalibrationForm**: Circle position detection calibration
- **ContrastCalibrationForm**: Contrast threshold adjustment
- **PixelCalibrationForm**: Per-pixel anomaly threshold
- **WhiteCalibrationForm**: White balance reference
- **ObjectBiasCalibrationForm**: Object position bias correction
- **gapThreshCalibrationForm**: Gap detection threshold

### Test Files (mostly legacy)
- `testAOI.cs`, `testAOI2.cs`, `testPerPixel.cs`, `testroi.cs`, `onnxTest.cs`, `onnx_Test.cs`
- These are development test harnesses, not part of main application flow

## Logging

### Serilog Configuration
- Main log: `.\logs\peilin_log-{Date}.txt`
- Camera-specific logs: `CameraLog_{Date}.txt`, `CameraWarning_{Date}.txt`
- Template: `{Timestamp:HH:mm:ss} [{Level:u3}] {Message}{NewLine}{Exception}`

## Multi-Camera Coordination

Each of the 4 cameras operates independently but shares:
1. **PLC communication channel** (serial port, Modbus protocol)
2. **Database connection pool** (SQLite with LinqToDB)
3. **Global `app` state** for cross-station coordination
4. **Shared disposal queue** (`_disposeQueue`) with semaphore limiting 4 concurrent disposals

The system uses:
- `_wh1` through `_wh4` wait handles for per-station image processing
- Concurrent queues to isolate station processing pipelines
- Station ID embedded in image metadata (`ImagePosition` class)

## Key Workflows

### Adding New Part Number
1. Use `type_info.cs` form to add part number to database
2. Configure parameters via `ParameterConfigForm`:
   - Select source part number as reference
   - Copy and modify parameters
   - Run calibration tools as needed
   - Save to database
3. Add AI models to `bin\x64\Release\models\{PartNumber}_*.pt`
4. Create YOLO batch file if object detection is required
5. Test with sample images before production

### Modifying Detection Parameters
1. Open `ParameterConfigForm`
2. Load target part number
3. Select source reference (can be same part number)
4. Parameters are categorized:
   - Camera: exposure, gain, delay
   - Position: circle detection, ROI definitions
   - Detection: anomaly thresholds, objBias, contrast
   - Timing: delays and timeouts
   - Testing: validation parameters
5. Use calibration tools for position/visual parameters
6. Save to database (only saves Added zones, preserves Reference)

### Debugging Image Processing Issues
1. Enable image saving in Parameters table
2. Check `bin\x64\Release\report\{PartNumber}\` for saved images
3. Review logs in `.\logs\` for timing and errors
4. Use Camera logs to diagnose grabbing issues
5. Check PLC communication logs for trigger signal problems

## Performance Considerations

- **Image disposal**: Use async disposal queue to prevent UI blocking
- **AI inference**: Each station can use different TensorRT DLL instance (1-4)
- **Database writes**: Batch when possible, use transactions for multi-row updates
- **PLC commands**: Priority queue prevents air blow command delays
- **YOLO warmup**: System auto-warms up YOLO models after periods of inactivity
- **Mat objects**: Never hold references longer than needed; disposal is critical

## Common Pitfalls

1. **Not disposing OpenCV Mat**: Causes rapid memory growth and crashes
2. **Using unqualified `Point`/`Size`**: Ambiguous between System.Drawing and OpenCvSharp
3. **Direct Insert without checking existence**: Causes primary key constraint violations
4. **Modifying `app` state without synchronization**: Race conditions across stations
5. **Blocking PLC serial port**: Air blow commands must be highest priority
6. **Forgetting to load models before inference**: Check model paths exist before CreateModel calls

## YOLO Server Management

The system uses external Python processes for YOLO detection:
- Start via `bin\x64\Release\yolo_{PartNumber}_{StationNum}.bat`
- Communicates via HTTP on localhost:5001-5004
- Supports model hot-swapping via `/load_model` endpoint
- Warmup mechanism to keep models in GPU memory
- Batch detection endpoint: `/detect_batch`

Example workflow (handled by `YoloDetection.cs`):
1. Start server via batch file
2. Wait for server availability
3. Load model for current part number
4. Warmup with dummy images
5. Send detection requests with Base64-encoded images
6. Parse JSON responses with bounding boxes and class IDs

## Documentation Management Rules

### Bilingual Documentation Requirement
**CRITICAL:** This project maintains documentation in both English and Traditional Chinese.

When generating or updating `CLAUDE.md` (English version), you **MUST** simultaneously generate or update `CLAUDE-zh-tw.md` (Traditional Chinese version):
- Both files must contain the same information, differing only in language
- `CLAUDE.md` must be in English
- `CLAUDE-zh-tw.md` must be in Traditional Chinese (complete translation of CLAUDE.md)
- Any structural changes, new sections, or content updates must be synchronized across both files

### Git Push Requirements
After making any documentation or code changes, follow this workflow:

1. **Stage and commit changes:**
   ```bash
   git add .
   git commit -m "..."
   git push
   ```

2. **Verify push success:**
   ```bash
   git status
   # or
   git log --oneline -3
   ```

### Structured Commit Message Format
Every file change must be explicitly documented in the commit message using this structure:

```
<type>(<scope>): <subject>

<body>
- File1: description of changes and rationale
- File2: description of changes and rationale
- File3: description of changes and rationale

<footer>
```

**Commit Message Requirements:**

- **Type:** feat, fix, docs, refactor, chore, test, style, etc.
- **Scope:** Affected module or component (e.g., ai, notion, webhook, docs)
- **Subject:** Brief summary (50 characters or less)
- **Body:** Detailed explanation for each changed file:
  - What changed in each file
  - Why this change was necessary
  - Impact of the change
- **Footer:** Co-authors, references, breaking changes

**Examples:**

```
docs(project): Add bilingual documentation management rules

- CLAUDE.md: Added documentation management section with bilingual requirements, git workflow, and commit message standards
- CLAUDE-zh-tw.md: Created complete Traditional Chinese translation of CLAUDE.md

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude <noreply@anthropic.com>
```

```
feat(parameters): Implement 3-zone parameter management system

- ParameterConfigForm.cs: Refactored UI to support Reference/Added-Unmodified/Added-Modified zones
- ParameterSetupManager.cs: Added zone tracking logic and save validation
- CLAUDE.md: Updated documentation to reflect new parameter management workflow
- CLAUDE-zh-tw.md: Synchronized Traditional Chinese documentation

This change improves parameter editing workflow by clearly distinguishing source parameters from edited copies, preventing accidental overwrites.

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude <noreply@anthropic.com>
```

### Documentation Synchronization
When making any code changes, you must review and update `CLAUDE.md` if needed.

**Key areas requiring documentation updates:**
- Architecture changes (new modules, modified flows)
- New features or commands
- Database schema modifications
- API integrations or external service changes
- Configuration or environment variable changes
- Workflow or business logic changes

**Synchronization workflow:**
1. Make code changes
2. Evaluate if changes impact documentation
3. If yes, update `CLAUDE.md` (English)
4. Immediately update `CLAUDE-zh-tw.md` (Traditional Chinese translation)
5. Commit both documentation files together with code changes

**Documentation must always reflect the current state of the codebase to prevent information drift.**
