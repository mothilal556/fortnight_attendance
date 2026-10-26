from vercel.blob import sync as blob_sync


def download_private_blob(pathname):
    return blob_sync.get_blob(pathname, access="private").body


def upload_private_blob(pathname, content, content_type):
    return blob_sync.put_blob(
        pathname,
        content,
        access="private",
        content_type=content_type,
        add_random_suffix=False,
    )


def delete_blob(pathname):
    blob_sync.delete_blob(pathname)
