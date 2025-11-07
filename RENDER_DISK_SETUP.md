# Render Persistent Disk Setup

**Yes! Render supports persistent storage natively** - no external services needed!

## Quick Setup (2 minutes)

### Step 1: Add Disk in Render Dashboard

1. Go to your service: https://dashboard.render.com/web/srv-d47459p5pdvs73dm4fa0
2. Click **"Settings"** in the left sidebar
3. Scroll down to **"Disks"** section
4. Click **"Add Disk"**
5. Configure:
   - **Name**: `documents` (or any name)
   - **Size**: Start with 1GB (you can increase later)
   - **Mount Path**: `/mnt/disk` (default - don't change this)
6. Click **"Save"**

### Step 2: Verify Environment Variables

The service is already configured! Check that these are set:
- `STORAGE_TYPE=disk` ✅ (already set)
- `DISK_PATH=/mnt/disk/documents` ✅ (already set)
- `BASE_URL=https://office-word-mcp.onrender.com` ✅ (already set)

### Step 3: Deploy

Render will automatically redeploy when you add the disk. Documents will now persist!

## How It Works

- **Documents stored at**: `/mnt/disk/documents/`
- **Persists across**: Restarts, redeployments, updates
- **Access via**: `https://office-word-mcp.onrender.com/documents/{filename}.docx`

## Important Notes

⚠️ **Single Instance Only**: Services with persistent disks cannot scale to multiple instances. This is fine for most use cases.

✅ **Zero Configuration**: Once the disk is attached, everything works automatically!

## Cost

- **Starter Plan**: Included (up to 1GB free)
- **Higher Plans**: ~$0.25/GB/month

## Testing

After adding the disk, test with:

```bash
# Create a document
curl -X POST https://office-word-mcp.onrender.com/mcp/stream \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": 1,
    "method": "tools/call",
    "params": {
      "name": "create_document",
      "arguments": {
        "filename": "test.docx",
        "title": "Test Document"
      }
    }
  }'
```

The document will persist even after the service restarts!

