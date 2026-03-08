"""OneDrive URL to local filesystem path resolver.

When a PowerPoint file is stored in OneDrive, the COM API's FullName property
returns an HTTP(S) URL (e.g. https://d.docs.live.net/<CID>/Documents/file.pptx)
instead of a local filesystem path. This module resolves such URLs back to the
actual local path by checking the Windows registry and environment variables.
"""

import logging
import os
import winreg
from typing import Optional
from urllib.parse import unquote

logger = logging.getLogger(__name__)


def resolve_local_path(full_name: str) -> Optional[str]:
    """Resolve a OneDrive URL to the corresponding local filesystem path.

    Args:
        full_name: The FullName property from a PowerPoint Presentation COM
            object. May be a local path or a OneDrive HTTP(S) URL.

    Returns:
        The local filesystem path if resolution succeeds, or the original
        full_name if it is already a local path. Returns None if the URL
        cannot be resolved to a local path.
    """
    try:
        # Not a URL — already a local path
        if not full_name.startswith("http"):
            return full_name

        # Try registry-based resolution first
        local_path = _resolve_via_registry(full_name)
        if local_path:
            return local_path

        # Fall back to environment variable-based resolution
        local_path = _resolve_via_env(full_name)
        if local_path:
            return local_path

        logger.debug("Could not resolve OneDrive URL to local path: %s", full_name)
        return None
    except Exception:
        logger.debug("Error resolving OneDrive URL: %s", full_name, exc_info=True)
        return None


def _resolve_via_registry(full_name: str) -> Optional[str]:
    """Try to resolve a OneDrive URL using the Windows registry.

    Enumerates HKCU\\Software\\SyncEngines\\Providers\\OneDrive subkeys.
    Each subkey has UrlNamespace (URL prefix) and MountPoint (local dir).
    """
    try:
        reg_path = r"Software\SyncEngines\Providers\OneDrive"
        try:
            providers_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path)
        except FileNotFoundError:
            return None

        try:
            i = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(providers_key, i)
                    i += 1
                except OSError:
                    break

                try:
                    subkey = winreg.OpenKey(providers_key, subkey_name)
                    try:
                        url_namespace, _ = winreg.QueryValueEx(
                            subkey, "UrlNamespace"
                        )
                        mount_point, _ = winreg.QueryValueEx(
                            subkey, "MountPoint"
                        )
                    finally:
                        winreg.CloseKey(subkey)

                    if not url_namespace or not mount_point:
                        continue

                    # Ensure URL namespace ends with / for proper prefix matching
                    if not url_namespace.endswith("/"):
                        url_namespace += "/"

                    if full_name.startswith(url_namespace):
                        relative = full_name[len(url_namespace):]
                        relative = unquote(relative)
                        relative = relative.replace("/", "\\")

                        # Ensure mount_point does not end with backslash
                        mount_point = mount_point.rstrip("\\")
                        local_path = mount_point + "\\" + relative
                        logger.debug(
                            "Resolved via registry: %s -> %s",
                            full_name,
                            local_path,
                        )
                        return local_path
                except (FileNotFoundError, OSError):
                    continue
        finally:
            winreg.CloseKey(providers_key)
    except Exception:
        logger.debug("Registry resolution failed", exc_info=True)

    return None


def _resolve_via_env(full_name: str) -> Optional[str]:
    """Try to resolve a OneDrive URL using environment variables.

    For personal OneDrive (https://d.docs.live.net/<CID>/...):
        Uses %OneDriveConsumer% or %OneDrive%
    For SharePoint/OneDrive for Business (*.sharepoint.com):
        Uses %OneDriveCommercial%
    """
    try:
        decoded_url = unquote(full_name)

        # Personal OneDrive: https://d.docs.live.net/<CID>/path/to/file
        if "d.docs.live.net" in decoded_url:
            # Extract path after the CID segment
            # URL format: https://d.docs.live.net/<CID>/rest/of/path
            parts = decoded_url.split("/")
            # Find the CID (comes after d.docs.live.net)
            try:
                host_idx = next(
                    i
                    for i, p in enumerate(parts)
                    if p == "d.docs.live.net"
                )
                # Path segments start after CID
                relative_parts = parts[host_idx + 2:]
            except StopIteration:
                return None

            if not relative_parts:
                return None

            relative = "\\".join(relative_parts)

            # Try OneDriveConsumer first, then OneDrive
            for env_var in ("OneDriveConsumer", "OneDrive"):
                mount = os.environ.get(env_var)
                if mount and os.path.isdir(mount):
                    local_path = os.path.join(mount, relative)
                    logger.debug(
                        "Resolved via %%%s%%: %s -> %s",
                        env_var,
                        full_name,
                        local_path,
                    )
                    return local_path

        # SharePoint / OneDrive for Business: https://<tenant>.sharepoint.com/...
        if ".sharepoint.com" in decoded_url:
            # Extract path after /Documents/ or /Shared Documents/
            for marker in ("/Documents/", "/Shared%20Documents/", "/Shared Documents/"):
                idx = decoded_url.find(marker)
                if idx >= 0:
                    relative = decoded_url[idx + len(marker):]
                    relative = relative.replace("/", "\\")

                    mount = os.environ.get("OneDriveCommercial")
                    if mount and os.path.isdir(mount):
                        local_path = os.path.join(mount, relative)
                        logger.debug(
                            "Resolved via %%OneDriveCommercial%%: %s -> %s",
                            full_name,
                            local_path,
                        )
                        return local_path
                    break
    except Exception:
        logger.debug("Environment variable resolution failed", exc_info=True)

    return None
