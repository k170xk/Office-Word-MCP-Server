# Storage Configuration Guide

The Office Word MCP Server supports multiple storage backends for document persistence. Choose the option that best fits your needs.

## Storage Options

### 1. **AWS S3 (Recommended for Production)**

Best for: Production deployments, scalability, multi-instance support

**Setup:**
1. Create an S3 bucket in AWS
2. Create an IAM user with S3 read/write permissions
3. Set environment variables in Render:

```bash
STORAGE_TYPE=s3
S3_BUCKET_NAME=your-bucket-name
S3_REGION=us-east-1
AWS_ACCESS_KEY_ID=your-access-key
AWS_SECRET_ACCESS_KEY=your-secret-key
BASE_URL=https://office-word-mcp.onrender.com
```

**Benefits:**
- ✅ Documents persist across deployments
- ✅ Works with multiple service instances
- ✅ Scalable and reliable
- ✅ Can enable versioning for document history
- ✅ Direct S3 URLs for documents

**Cost:** ~$0.023 per GB/month + transfer costs

---

### 2. **Render Disk (Good for MVP/Development)**

Best for: Simple setup, single-instance deployments, development

**Setup:**
1. In Render dashboard, go to your service
2. Add a Disk volume (Settings → Disks)
3. Mount it to `/mnt/disk`
4. Set environment variables:

```bash
STORAGE_TYPE=disk
DISK_PATH=/mnt/disk/documents
BASE_URL=https://office-word-mcp.onrender.com
```

**Benefits:**
- ✅ Simple setup
- ✅ Documents persist across restarts
- ✅ No additional service needed
- ✅ Fast local access

**Limitations:**
- ❌ Tied to single instance
- ❌ Lost if service is deleted
- ❌ Not suitable for scaling

**Cost:** Included in Render plan (varies by plan)

---

### 3. **Local Storage (Default - Ephemeral)**

Best for: Testing only, not recommended for production

**Setup:**
```bash
STORAGE_TYPE=local
DOCUMENTS_DIR=./documents
BASE_URL=https://office-word-mcp.onrender.com
```

**Warning:** Documents are lost when container restarts!

---

## How It Works

1. **Document Creation/Editing:**
   - Document is downloaded from storage to temp location
   - Tool operates on local copy
   - Document is uploaded back to storage after modification
   - URL is returned for download

2. **Document Access:**
   - `/documents/{filename}` endpoint downloads from storage
   - Returns document for download

3. **Automatic Sync:**
   - All operations automatically sync with storage
   - No manual upload/download needed

## Example Workflow

```python
# Create document
create_document(filename="report.docx", title="Monthly Report")
# → Document saved to S3/Disk
# → Returns: "Document report.docx created successfully
#            Document URL: https://office-word-mcp.onrender.com/documents/report.docx"

# Edit document
add_paragraph(filename="report.docx", text="New content")
# → Downloads from storage
# → Adds paragraph
# → Uploads back to storage
# → Returns URL

# Download document
# GET https://office-word-mcp.onrender.com/documents/report.docx
# → Downloads from storage and serves
```

## Migration Between Storage Types

Documents are stored by filename only. To migrate:

1. List all documents: `list_available_documents()`
2. Download each document
3. Change `STORAGE_TYPE` environment variable
4. Documents will be created in new storage on next edit

## Recommended Setup for Production

```bash
# Use S3 for production
STORAGE_TYPE=s3
S3_BUCKET_NAME=my-word-documents
S3_REGION=us-east-1
AWS_ACCESS_KEY_ID=AKIA...
AWS_SECRET_ACCESS_KEY=...
BASE_URL=https://office-word-mcp.onrender.com

# Optional: Enable S3 versioning for document history
# Configure in AWS S3 console → Bucket → Versioning
```

## Troubleshooting

**Documents not persisting:**
- Check `STORAGE_TYPE` is set correctly
- Verify S3 credentials (if using S3)
- Check Render Disk is mounted (if using disk)

**Slow operations:**
- S3 operations have network latency
- Consider using Render Disk for faster local access
- Or implement caching layer

**Permission errors:**
- S3: Check IAM user has `s3:GetObject`, `s3:PutObject`, `s3:DeleteObject` permissions
- Disk: Check disk is mounted and writable

