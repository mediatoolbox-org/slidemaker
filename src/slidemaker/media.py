"""
title: Media helpers for resolving image sources.
summary: |-
  Provides local-path resolution for image specifications, including
  remote URL downloads and prompt-based generation through pluggable
  providers such as OpenAI.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
from pathlib import Path
from typing import Any, Optional, Protocol
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen as _stdlib_urlopen

_IMAGE_EXTENSIONS = {
    ".bmp",
    ".gif",
    ".jpeg",
    ".jpg",
    ".png",
    ".tif",
    ".tiff",
    ".webp",
}
_CONTENT_TYPE_EXTENSIONS = {
    "image/bmp": ".bmp",
    "image/gif": ".gif",
    "image/jpeg": ".jpeg",
    "image/png": ".png",
    "image/tiff": ".tiff",
    "image/webp": ".webp",
}
_PROMPT_OPTION_KEYS = (
    "size",
    "quality",
    "background",
    "output_format",
    "output_compression",
    "moderation",
    "user",
)


class ImagePromptProvider(Protocol):
    """
    title: Protocol for prompt-based image providers.
    attributes:
      cache_key:
        type: str
        description: Stable identifier used when caching generated images.
    """

    cache_key: str

    def generate(
        self,
        prompt: str,
        *,
        options: Optional[dict[str, Any]] = None,
    ) -> bytes:
        """
        title: Generate one image for the given prompt.
        parameters:
          prompt:
            type: str
          options:
            type: Optional[dict[str, Any]]
        returns:
          type: bytes
        """


def _spec_value(spec: dict[str, Any], *keys: str, default: Any = None) -> Any:
    """
    title: Read a spec option supporting hyphen and underscore spellings.
    parameters:
      spec:
        type: dict[str, Any]
      default:
        type: Any
      keys:
        type: str
        variadic: positional
    returns:
      type: Any
    """
    for key in keys:
        if key in spec:
            return spec[key]
        alternate = key.replace("-", "_") if "-" in key else key.replace("_", "-")
        if alternate in spec:
            return spec[alternate]
    return default


def _content_type_extension(content_type: Optional[str]) -> Optional[str]:
    """
    title: Map a response content type to a file extension.
    parameters:
      content_type:
        type: Optional[str]
    returns:
      type: Optional[str]
    """
    if not content_type:
        return None
    normalized = content_type.split(";", 1)[0].strip().lower()
    return _CONTENT_TYPE_EXTENSIONS.get(normalized)


def _url_extension(url: str) -> Optional[str]:
    """
    title: Infer a likely image extension from a URL path.
    parameters:
      url:
        type: str
    returns:
      type: Optional[str]
    """
    suffix = Path(urlparse(url).path).suffix.lower()
    if suffix in _IMAGE_EXTENSIONS:
        return suffix
    return None


def _sha256_hex(value: str) -> str:
    """
    title: Compute a hex SHA-256 digest.
    parameters:
      value:
        type: str
    returns:
      type: str
    """
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


class OpenAIImageProvider:
    """
    title: Prompt-based image generation using OpenAI's Images API.
    attributes:
      _model:
        description: The OpenAI image model name.
      _api_key:
        description: Secret API key used for authentication.
      _base_url:
        description: Base URL for the OpenAI API.
      _timeout:
        description: Network timeout in seconds.
      _urlopen:
        description: Injectable opener used for HTTP requests.
      cache_key:
        description: Stable identifier used when caching generated images.
    """

    def __init__(
        self,
        model: str = "gpt-image-1",
        api_key: Optional[str] = None,
        api_key_env: str = "OPENAI_API_KEY",
        base_url: str = "https://api.openai.com/v1",
        timeout: float = 120.0,
        urlopen: Any = _stdlib_urlopen,
    ) -> None:
        """
        title: Configure the OpenAI image provider.
        parameters:
          model:
            type: str
          api_key:
            type: Optional[str]
          api_key_env:
            type: str
          base_url:
            type: str
          timeout:
            type: float
          urlopen:
            type: Any
        """
        resolved_api_key = api_key or os.environ.get(api_key_env)
        if not resolved_api_key:
            raise ValueError("OpenAI API key is required")
        self._model = str(model)
        self._api_key = resolved_api_key
        self._base_url = base_url.rstrip("/")
        self._timeout = float(timeout)
        self._urlopen = urlopen
        self.cache_key = f"openai:{self._base_url}:{self._model}"

    def generate(
        self,
        prompt: str,
        *,
        options: Optional[dict[str, Any]] = None,
    ) -> bytes:
        """
        title: Generate one image and return its raw bytes.
        parameters:
          prompt:
            type: str
          options:
            type: Optional[dict[str, Any]]
        returns:
          type: bytes
        """
        payload: dict[str, Any] = {
            "model": self._model,
            "prompt": prompt,
            "n": 1,
        }
        for key, value in (options or {}).items():
            if value is not None:
                payload[key] = value

        request = Request(
            f"{self._base_url}/images/generations",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self._api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )

        try:
            with self._urlopen(request, timeout=self._timeout) as response:
                response_bytes = response.read()
        except HTTPError as exc:
            error_body = exc.read().decode("utf-8", "replace")
            raise RuntimeError(
                f"OpenAI image generation failed: {exc.code} {error_body}".strip()
            ) from exc
        except URLError as exc:
            raise RuntimeError(f"OpenAI image generation failed: {exc.reason}") from exc

        try:
            payload_obj = json.loads(response_bytes.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise RuntimeError("OpenAI image generation returned invalid JSON") from exc

        data_items = payload_obj.get("data")
        if not isinstance(data_items, list) or not data_items:
            raise RuntimeError("OpenAI image generation returned no image data")

        first_item = data_items[0]
        if not isinstance(first_item, dict):
            raise RuntimeError(
                "OpenAI image generation returned an unexpected response"
            )

        b64_json = first_item.get("b64_json")
        if isinstance(b64_json, str):
            try:
                return base64.b64decode(b64_json)
            except (ValueError, TypeError) as exc:
                raise RuntimeError(
                    "OpenAI image generation returned invalid base64 data"
                ) from exc

        image_url = first_item.get("url")
        if isinstance(image_url, str):
            return self._download_response_bytes(image_url)

        raise RuntimeError("OpenAI image generation returned no usable image payload")

    def _download_response_bytes(self, url: str) -> bytes:
        """
        title: Download binary response bytes from a URL.
        parameters:
          url:
            type: str
        returns:
          type: bytes
        """
        request = Request(url, headers={"User-Agent": "slidemaker"})
        try:
            with self._urlopen(request, timeout=self._timeout) as response:
                return response.read()
        except HTTPError as exc:
            raise RuntimeError(f"OpenAI image download failed: {exc.code}") from exc
        except URLError as exc:
            raise RuntimeError(f"OpenAI image download failed: {exc.reason}") from exc


def build_image_provider(
    image_provider: ImagePromptProvider | dict[str, Any] | None,
) -> ImagePromptProvider | None:
    """
    title: Normalize image-provider configuration.
    parameters:
      image_provider:
        type: ImagePromptProvider | dict[str, Any] | None
    returns:
      type: ImagePromptProvider | None
    """
    if image_provider is None:
        return image_provider
    if not isinstance(image_provider, dict):
        if not hasattr(image_provider, "generate"):
            raise TypeError(
                "image_provider must be a provider object or configuration dict"
            )
        return image_provider

    provider_name = str(image_provider.get("provider", "openai")).strip().lower()
    if provider_name != "openai":
        raise ValueError(f"unknown image provider: {provider_name}")

    provider_config = dict(image_provider)
    provider_config.pop("provider", None)
    return OpenAIImageProvider(**provider_config)


class MediaResolver:
    """
    title: Resolve user-facing image specs into local image files.
    attributes:
      _image_provider:
        description: Provider used for prompt-based image generation.
      _cache_dir:
        description: >-
          Directory where downloaded and generated images are cached.
      _timeout:
        description: Network timeout in seconds.
      _urlopen:
        description: Injectable opener used for HTTP downloads.
    """

    def __init__(
        self,
        *,
        image_provider: ImagePromptProvider | None = None,
        cache_dir: str | Path | None = None,
        timeout: float = 60.0,
        urlopen: Any = _stdlib_urlopen,
    ) -> None:
        """
        title: Configure the media resolver.
        parameters:
          image_provider:
            type: ImagePromptProvider | None
          cache_dir:
            type: str | Path | None
          timeout:
            type: float
          urlopen:
            type: Any
        """
        self._image_provider = image_provider
        self._cache_dir = (
            Path(".slidemaker-cache") if cache_dir is None else Path(cache_dir)
        )
        self._timeout = float(timeout)
        self._urlopen = urlopen

    def resolve_image(
        self,
        image: str | Path | dict[str, Any] | None,
    ) -> str | Path | dict[str, Any] | None:
        """
        title: Resolve an image spec to a local-path image spec.
        parameters:
          image:
            type: str | Path | dict[str, Any] | None
        returns:
          type: str | Path | dict[str, Any] | None
        """
        if image is None or isinstance(image, (str, Path)):
            return image
        if not isinstance(image, dict):
            return image

        local_source = _spec_value(image, "path", "src")
        remote_url = _spec_value(image, "url")
        prompt = _spec_value(image, "prompt")
        source_count = sum(
            value is not None for value in (local_source, remote_url, prompt)
        )
        if source_count != 1:
            raise ValueError(
                "image spec must define exactly one of path, src, url, or prompt"
            )

        if local_source is not None:
            return image

        resolved = dict(image)
        if remote_url is not None:
            if not isinstance(remote_url, str):
                raise TypeError("image url must be a string")
            resolved["path"] = self._download_image(remote_url)
            resolved.pop("url", None)
            return resolved

        if not isinstance(prompt, str):
            raise TypeError("image prompt must be a string")
        resolved["path"] = self._generate_prompt_image(
            prompt,
            options=self._prompt_options(image),
        )
        resolved.pop("prompt", None)
        return resolved

    def _prompt_options(self, image_spec: dict[str, Any]) -> dict[str, Any]:
        """
        title: Extract provider-specific prompt options from an image spec.
        parameters:
          image_spec:
            type: dict[str, Any]
        returns:
          type: dict[str, Any]
        """
        options: dict[str, Any] = {}
        for key in _PROMPT_OPTION_KEYS:
            value = _spec_value(image_spec, key)
            if value is None:
                continue
            if key == "output_compression":
                if not isinstance(value, int):
                    raise TypeError("image output_compression must be an integer")
            elif not isinstance(value, str):
                raise TypeError(f"image {key} must be a string")
            options[key] = value
        return options

    def _download_image(self, url: str) -> Path:
        """
        title: Download an image URL into the cache directory.
        parameters:
          url:
            type: str
        returns:
          type: Path
        """
        cache_key = _sha256_hex(f"url:{url}")
        extension = _url_extension(url) or ".img"
        cache_path = self._cache_path("url", cache_key, extension)
        if cache_path.exists():
            return cache_path

        request = Request(url, headers={"User-Agent": "slidemaker"})
        try:
            with self._urlopen(request, timeout=self._timeout) as response:
                payload = response.read()
                response_headers = getattr(response, "headers", None)
        except HTTPError as exc:
            raise RuntimeError(f"image download failed: {exc.code}") from exc
        except URLError as exc:
            raise RuntimeError(f"image download failed: {exc.reason}") from exc

        content_type = None
        if response_headers is not None:
            if hasattr(response_headers, "get_content_type"):
                content_type = response_headers.get_content_type()
            elif hasattr(response_headers, "get"):
                content_type = response_headers.get("Content-Type")

        content_extension = _content_type_extension(content_type)
        if content_extension is not None and content_extension != cache_path.suffix:
            cache_path = self._cache_path("url", cache_key, content_extension)
            if cache_path.exists():
                return cache_path

        self._write_bytes(cache_path, payload)
        return cache_path

    def _generate_prompt_image(
        self,
        prompt: str,
        *,
        options: Optional[dict[str, Any]] = None,
    ) -> Path:
        """
        title: Generate a prompt-based image into the cache directory.
        parameters:
          prompt:
            type: str
          options:
            type: Optional[dict[str, Any]]
        returns:
          type: Path
        """
        if self._image_provider is None:
            raise ValueError("image prompt requires an image_provider")

        normalized_options = dict(options or {})
        output_format = str(normalized_options.get("output_format", "png")).lower()
        cache_material = json.dumps(
            {
                "provider": getattr(
                    self._image_provider,
                    "cache_key",
                    self._image_provider.__class__.__name__,
                ),
                "prompt": prompt,
                "options": normalized_options,
            },
            sort_keys=True,
        )
        cache_path = self._cache_path(
            "prompt",
            _sha256_hex(cache_material),
            f".{output_format}",
        )
        if cache_path.exists():
            return cache_path

        payload = self._image_provider.generate(prompt, options=normalized_options)
        self._write_bytes(cache_path, payload)
        return cache_path

    def _cache_path(self, prefix: str, digest: str, extension: str) -> Path:
        """
        title: Build a cache path for a resolved media artifact.
        parameters:
          prefix:
            type: str
          digest:
            type: str
          extension:
            type: str
        returns:
          type: Path
        """
        self._cache_dir.mkdir(parents=True, exist_ok=True)
        return self._cache_dir / f"{prefix}-{digest}{extension}"

    def _write_bytes(self, path: Path, payload: bytes) -> None:
        """
        title: Atomically write bytes to a cache file.
        parameters:
          path:
            type: Path
          payload:
            type: bytes
        """
        temp_path = path.with_suffix(f"{path.suffix}.tmp")
        temp_path.write_bytes(payload)
        temp_path.replace(path)
