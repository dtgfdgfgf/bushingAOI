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

### Minimal Code Changes Principle
**Critical:** When modifying or adding code, follow the principle of minimal changes.

**Philosophy:**
- Make the smallest possible changes to achieve the required functionality
- Preserve existing code structure and patterns whenever possible
- Only modify what is necessary to fix bugs or add features
- Respect the existing architecture and design decisions

**This does not mean compromising on code quality or architecture:**
- Maintain proper error handling, memory management, and thread safety
- Follow all other guidelines in this document
- Refactor when absolutely necessary for correctness or critical maintainability
- Design new features with proper architecture from the start

**Practical application:**
- **Bug fixes**: Change only the specific lines causing the issue
- **New features**: Integrate into existing patterns rather than rewriting surrounding code
- **Parameter additions**: Add new fields without restructuring existing data flow
- **Performance improvements**: Optimize specific bottlenecks, not entire subsystems
- **Refactoring**: Only when code is broken, unmaintainable, or creates production risks

**Example - Adding a new parameter:**
```csharp
// GOOD: Minimal change - add new parameter alongside existing ones
if (updatedRows == 0)
{
    db.params.Value(p => p.PartNumber, partNum)
             .Value(p => p.ExistingParam, existingVal)
             .Value(p => p.NewParam, newVal)  // Only this line added
             .Insert();
}

// AVOID: Unnecessary restructuring of working code
// Don't rewrite the entire parameter loading system just to add one field
```

**Balance:** The goal is stability and predictability in a production system. Make surgical changes that solve the problem without introducing new risks. However, never sacrifice critical requirements (memory safety, thread safety, data integrity) for the sake of minimal changes.

### Code Change Communication
**When proposing or making code changes, always provide:**

1. **Precise location**: File path, class name, method name, and line number/range
2. **Clear rationale**: Why the change is necessary (bug fix, new feature, performance, safety)
3. **Change scope**: What will be modified and what will remain unchanged
4. **Impact assessment**: Which components/threads/workflows will be affected

**Example format:**
```
Location: Form1.cs:1234 in ProcessImage() method
Reason: Memory leak - Mat object not disposed after YOLO inference
Change: Add using statement around Mat creation (lines 1234-1240)
Impact: Only affects YOLO processing path; no impact on anomaly detection or PLC communication
```

**This ensures:**
- Clear understanding of what will change before implementation
- Minimal unexpected modifications
- Better review and debugging capability
- Documentation of change history

### Language & Comments
- **All comments must be in Traditional Chinese (繁體中文)**
- Add comment at start of AI-generated code sections: `// 由 GitHub Copilot 產生`

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

#### Strict Memory Discipline
- **OpenCV Mat objects**: Must be disposed immediately after use
- **System.Drawing.Bitmap**: Must call `.Dispose()` explicitly
- Use `using` statements for image processing operations
- Long-lived large objects should use compressed storage
- The application uses `_disposeQueue` and `_disposeSemaphore` for async disposal

#### Memory Overload Prevention
**This system operates under extreme memory constraints due to high-resolution image processing at industrial speeds. Memory overload leads to application crashes and production line failures.**

**Mandatory practices:**
1. **Immediate disposal**: Dispose Mat/Bitmap objects within the same method scope whenever possible
2. **Queue management**: Monitor queue depths (`Queue_Bitmap1-4`, `Queue_Save`, `Queue_Send`, `Queue_Show`) to prevent memory buildup
3. **Large object lifecycle**: Track and limit the lifetime of objects >85KB (Large Object Heap threshold)
4. **Avoid GC.Collect()**: Never call `GC.Collect()` explicitly unless absolutely necessary and documented; let the runtime manage collection
5. **Image copy minimization**: Avoid unnecessary image cloning; use references where safe, dispose originals when cloned
6. **Memory profiling**: Regularly profile memory usage under production load conditions

### Multi-Threading Management
**Critical:** This application runs 4+ concurrent processing pipelines with shared resources. Thread safety violations cause data corruption and crashes.

#### Thread Safety Requirements
**The system coordinates multiple threads processing camera feeds, AI inference, PLC communication, and UI updates. Improper synchronization causes:**
- Data races on shared state
- Resource leaks (unreleased locks, undisposed objects)
- Cache inconsistency across CPU cores
- Deadlocks in PLC communication

**Mandatory practices:**

1. **Synchronization primitives**:
   - Use `lock` statements for all `app` class static state modifications
   - Use wait handles (`_wh1-4`, `_sv`, `_reader`, `_AI`, `_show`) correctly with timeout patterns
   - Use `SemaphoreSlim` for async operations (e.g., `_disposeSemaphore` limits concurrent disposals)

2. **Thread-safe collections**:
   - Use `ConcurrentQueue<T>` for all cross-thread queues (`Queue_Bitmap1-4`, etc.)
   - Never iterate concurrent collections without understanding snapshot semantics
   - Use `TryDequeue`/`TryPeek` patterns, never direct indexing

3. **Shared resource access**:
   - **PLC SerialPort**: Single writer pattern with priority queue; never block the port
   - **Database connections**: Always use separate connection per thread via `using` blocks; never share connection objects
   - **Camera objects**: Each camera instance is accessed only by its dedicated thread; never cross-reference
   - **TensorRT DLL instances**: Each station uses a dedicated instance (1-4); never share inference sessions

4. **Data race prevention**:
   - **Read-modify-write operations**: Always lock before reading shared counters/flags if the value determines subsequent writes
   - **Event handlers**: Assume events fire on arbitrary threads; marshal to UI thread for control updates via `Invoke`/`BeginInvoke`
   - **Static caches**: Use `ConcurrentDictionary` for shared lookup tables; never use plain `Dictionary` for cross-thread access

5. **Memory visibility**:
   - Use `volatile` for simple flags read/written across threads without locks
   - Use `Interlocked` operations for atomic counter updates
   - Understand that locks provide memory barriers; unsynchronized reads may see stale cached values

6. **Deadlock prevention**:
   - **Lock ordering**: Always acquire locks in consistent order (document the order if multiple locks needed)
   - **Timeout patterns**: Use `Monitor.TryEnter` or wait handle timeouts; never indefinite blocking
   - **Avoid nested locks**: Minimize lock scope; never call external code while holding a lock

7. **Resource cleanup in multi-threaded contexts**:
   - Ensure disposal queue (`_disposeQueue`) is drained on shutdown
   - Use cancellation tokens for graceful thread termination
   - Join worker threads before disposing shared resources

**Common thread safety violations to avoid:**
- Modifying `app.counter`, `app.param`, `app.models` without locks
- Sharing `DataConnection` objects across threads
- Calling `SerialPort.Write` from multiple threads without serialization
- Accessing Windows Forms controls from non-UI threads
- Disposing objects still referenced by worker threads
- Creating race conditions between PLC trigger events and image processing state

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

**All modifications to `app` state must be synchronized. Use locks for compound operations.**

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

**All shared resource access must be properly synchronized to prevent race conditions.**

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
- **Memory pressure**: Monitor queue depths and image object lifecycles under sustained load
- **Thread contention**: Profile lock contention on `app` state during peak throughput

## Common Pitfalls

1. **Not disposing OpenCV Mat**: Causes rapid memory growth and crashes
2. **Using unqualified `Point`/`Size`**: Ambiguous between System.Drawing and OpenCvSharp
3. **Direct Insert without checking existence**: Causes primary key constraint violations
4. **Modifying `app` state without synchronization**: Race conditions across stations
5. **Blocking PLC serial port**: Air blow commands must be highest priority
6. **Forgetting to load models before inference**: Check model paths exist before CreateModel calls
7. **Sharing database connections across threads**: Always create new connection per thread
8. **Calling GC.Collect() excessively**: Degrades performance; removed in recent commits
9. **Holding locks while calling external code**: Risk of deadlock; minimize lock scope
10. **Ignoring queue depth monitoring**: Uncontrolled queue growth leads to memory exhaustion

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
