# AI Training Dataset Management Recommendations

## Current System Analysis

Your current dataset structure in `bin/x64/Release/image/` shows excellent organization with:
- Date-based hierarchy (YYYY-MM/MMDD)
- Product type identification
- Timestamp tracking
- Defect classification with confidence scores
- Station-based organization
- ROI separation

## Recommended Enhancements

### 1. Standardized Naming Convention

**Current**: Mixed naming patterns across different dates
**Recommended**: Consistent pattern across all folders

```
YYYY-MM/
└── MMDD(status)/
    └── {ProductType}-{HHMM}({defect_summary})/
        ├── metadata.json
        ├── Stations/
        │   ├── Station{N}/
        │   │   └── {seq}-{station}-{defect}-{confidence}.{ext}
        ├── ROI_{N}/
        ├── origin/
        ├── processed/
        └── annotations/
```

### 2. Metadata Management

Create JSON files for each dataset session:

```json
{
  "session_id": "20250528_10214090420T_1644",
  "product_type": "10214090420T",
  "date": "2025-05-28",
  "time": "16:44",
  "status": "processed",
  "total_images": 240,
  "defect_summary": {
    "OK": 120,
    "dirty": 85,
    "bp": 20,
    "wr": 15
  },
  "stations": [1, 2, 3, 4],
  "roi_count": 4,
  "training_used": false,
  "model_version": null,
  "operator": "user_name",
  "notes": "Initial collection batch"
}
```

### 3. Training Dataset Organization

Create a separate training structure:

```
AI_Training_Datasets/
├── Projects/
│   ├── Project_001_DefectDetection_v1.0/
│   │   ├── config.json
│   │   ├── datasets/
│   │   │   ├── train/
│   │   │   ├── validation/
│   │   │   └── test/
│   │   ├── annotations/
│   │   └── models/
│   └── Project_002_MultiStation_v2.0/
├── Source_Mappings/
│   └── dataset_sources.json
└── Version_Control/
    └── training_history.json
```

### 4. Database Integration Schema

Extend your existing database with training-specific tables:

```sql
-- Training Projects Table
CREATE TABLE training_projects (
    project_id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_name TEXT NOT NULL,
    version TEXT NOT NULL,
    created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
    status TEXT DEFAULT 'active',
    description TEXT,
    model_type TEXT,
    target_defects TEXT -- JSON array of defect types
);

-- Dataset Sources Table  
CREATE TABLE dataset_sources (
    source_id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER,
    source_path TEXT NOT NULL,
    date_collected DATE,
    product_type TEXT,
    station_id INTEGER,
    image_count INTEGER,
    defect_counts TEXT, -- JSON object with defect counts
    inclusion_date DATETIME,
    FOREIGN KEY (project_id) REFERENCES training_projects(project_id)
);

-- Training History Table
CREATE TABLE training_history (
    training_id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER,
    training_date DATETIME DEFAULT CURRENT_TIMESTAMP,
    dataset_hash TEXT, -- Hash of dataset composition
    model_path TEXT,
    performance_metrics TEXT, -- JSON with accuracy, precision, recall
    validation_results TEXT, -- JSON with validation details
    notes TEXT,
    FOREIGN KEY (project_id) REFERENCES training_projects(project_id)
);
```

### 5. Automated Tools Recommendations

#### A. Dataset Analyzer Script
```powershell
# PowerShell script to analyze current dataset
$imagePath = "bin\x64\Release\image"
$analysisReport = @{}

# Scan and categorize all images
Get-ChildItem -Path $imagePath -Recurse -Include *.jpg,*.png | ForEach-Object {
    # Extract metadata from filename and path
    # Generate summary statistics
    # Identify potential training candidates
}
```

#### B. Training Preparation Tool
```csharp
// C# tool to prepare training datasets
public class TrainingDatasetManager
{
    public void CreateTrainingSet(string sourceDir, string targetDir, 
                                TrainingConfig config)
    {
        // Copy and organize images for training
        // Generate annotation files
        // Create train/validation/test splits
        // Update database records
    }
}
```

### 6. Quality Control Checklist

**Before Training:**
- [ ] Verify image quality and consistency
- [ ] Check defect label accuracy
- [ ] Ensure balanced dataset across defect types
- [ ] Validate station-specific requirements
- [ ] Review confidence score distributions

**During Training:**
- [ ] Monitor training metrics
- [ ] Log dataset composition
- [ ] Track model versions
- [ ] Document parameter changes

**After Training:**
- [ ] Validate on test data
- [ ] Update database with results
- [ ] Archive training artifacts
- [ ] Document deployment decisions

### 7. File Naming Standards

**For Raw Data:**
`{sequence}-{station}-{defect_type}-{confidence}.{extension}`
Example: `0001-1-dirty-0.82.png`

**For Processed Data:**
`{product_type}_{date}_{time}_{sequence}_{station}_{defect}_{confidence}.{extension}`
Example: `10214090420T_20250528_1644_0001_1_dirty_0.82.png`

**For Training Sets:**
`{split}_{product_type}_{defect_type}_{sequence}.{extension}`
Example: `train_10214090420T_dirty_0001.png`

### 8. Backup and Version Control Strategy

**Recommended Structure:**
```
Backups/
├── Raw_Data_Backups/
│   └── YYYY-MM/
├── Training_Snapshots/
│   └── Project_{ID}_v{version}/
└── Model_Archives/
    └── {model_name}_v{version}/
```

**Git Integration:**
- Use Git LFS for large image files
- Version control metadata and configuration files
- Tag releases with model versions
- Maintain change logs

### 9. Performance Monitoring

Track these metrics in your database:
- Dataset size vs. model performance
- Training time per dataset size
- Defect detection accuracy by station
- False positive/negative rates
- Model deployment success rates

### 10. Next Steps

1. **Immediate Actions:**
   - Standardize folder naming for new collections
   - Create metadata.json files for existing datasets
   - Implement database extensions

2. **Short-term (1-2 weeks):**
   - Develop automated dataset analysis tools
   - Create training preparation workflows
   - Establish backup procedures

3. **Medium-term (1 month):**
   - Implement full training pipeline integration
   - Deploy quality control checkpoints
   - Create performance dashboards

4. **Long-term (3 months):**
   - Automate entire workflow
   - Implement continuous learning pipeline
   - Establish model update procedures
