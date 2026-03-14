from __future__ import annotations

import base64
import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from urllib.error import HTTPError, URLError

import pytest
from slidemaker import OpenAIImageProvider
from slidemaker.media import (
    MediaResolver,
    _content_type_extension,
    _spec_value,
    _url_extension,
    build_image_provider,
)


class _FakeHeaders:
    def __init__(self, content_type: str) -> None:
        self._content_type = content_type

    def get_content_type(self) -> str:
        return self._content_type


class _FakeResponse:
    def __init__(self, payload: bytes, content_type: str = "application/json") -> None:
        self._payload = payload
        self.headers = _FakeHeaders(content_type)

    def read(self) -> bytes:
        return self._payload

    def __enter__(self) -> _FakeResponse:
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        return None


class _FakeHeadersMapping(dict):
    pass


class _FakePromptProvider:
    cache_key = "fake-provider:test"

    def __init__(self) -> None:
        self.calls: list[tuple[str, dict[str, object]]] = []

    def generate(
        self,
        prompt: str,
        *,
        options: dict[str, object] | None = None,
    ) -> bytes:
        self.calls.append((prompt, dict(options or {})))
        return b"generated-image"


class MediaTests(unittest.TestCase):
    def test_media_helper_functions_cover_fallback_cases(self) -> None:
        self.assertEqual(_spec_value({"output_format": "png"}, "output-format"), "png")
        self.assertEqual(_spec_value({"output-format": "png"}, "output_format"), "png")
        self.assertEqual(_spec_value({}, "missing", default="fallback"), "fallback")
        self.assertIsNone(_content_type_extension(None))
        self.assertEqual(_content_type_extension("image/png; charset=binary"), ".png")
        self.assertIsNone(_url_extension("https://example.com/assets/no-extension"))

    def test_build_image_provider_supports_openai_dict_config(self) -> None:
        provider = build_image_provider(
            {
                "provider": "openai",
                "model": "gpt-image-1",
                "api_key": "test-key",
            }
        )
        self.assertIsInstance(provider, OpenAIImageProvider)

        with self.assertRaisesRegex(ValueError, "unknown image provider"):
            build_image_provider({"provider": "bad", "api_key": "test-key"})

        with self.assertRaisesRegex(
            TypeError, "image_provider must be a provider object or configuration dict"
        ):
            build_image_provider(123)  # type: ignore[arg-type]

        prompt_provider = _FakePromptProvider()
        self.assertIs(build_image_provider(prompt_provider), prompt_provider)

    def test_openai_image_provider_builds_request_and_decodes_b64(self) -> None:
        calls: list[tuple[object, float]] = []

        def fake_urlopen(request: object, timeout: float) -> _FakeResponse:
            calls.append((request, timeout))
            payload = json.dumps(
                {"data": [{"b64_json": base64.b64encode(b"png-bytes").decode("ascii")}]}
            ).encode("utf-8")
            return _FakeResponse(payload)

        provider = OpenAIImageProvider(
            api_key="test-key",
            model="gpt-image-1",
            urlopen=fake_urlopen,
        )
        result = provider.generate("red fox", options={"size": "1024x1024"})

        self.assertEqual(result, b"png-bytes")
        request, timeout = calls[0]
        self.assertEqual(timeout, 120.0)
        self.assertEqual(
            request.full_url,
            "https://api.openai.com/v1/images/generations",
        )
        self.assertEqual(request.get_method(), "POST")
        self.assertEqual(request.headers["Authorization"], "Bearer test-key")
        body = json.loads(request.data.decode("utf-8"))
        self.assertEqual(body["model"], "gpt-image-1")
        self.assertEqual(body["prompt"], "red fox")
        self.assertEqual(body["size"], "1024x1024")

    def test_openai_image_provider_init_requires_api_key(self) -> None:
        with pytest.raises(ValueError, match="OpenAI API key is required"):
            OpenAIImageProvider(api_key=None, api_key_env="MISSING_TEST_KEY")

    def test_openai_image_provider_uses_generated_image_url(self) -> None:
        calls: list[tuple[object, float]] = []

        def fake_urlopen(request: object, timeout: float) -> _FakeResponse:
            calls.append((request, timeout))
            if len(calls) == 1:
                payload = json.dumps(
                    {"data": [{"url": "https://example.com/generated.png"}]}
                ).encode("utf-8")
                return _FakeResponse(payload)
            return _FakeResponse(b"downloaded-image", content_type="image/png")

        provider = OpenAIImageProvider(api_key="test-key", urlopen=fake_urlopen)
        result = provider.generate("fox")

        self.assertEqual(result, b"downloaded-image")
        self.assertEqual(calls[1][0].full_url, "https://example.com/generated.png")

    def test_openai_image_provider_reports_request_and_payload_errors(self) -> None:
        provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: (_ for _ in ()).throw(URLError("offline")),
        )
        with self.assertRaisesRegex(
            RuntimeError, "OpenAI image generation failed: offline"
        ):
            provider.generate("fox")

        http_error = HTTPError(
            url="https://api.openai.com/v1/images/generations",
            code=400,
            msg="Bad Request",
            hdrs=None,
            fp=None,
        )
        http_error.read = lambda: b'{"error":"bad"}'  # type: ignore[assignment]
        provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: (_ for _ in ()).throw(http_error),
        )
        with self.assertRaisesRegex(
            RuntimeError, "OpenAI image generation failed: 400"
        ):
            provider.generate("fox")

        bad_json_provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: _FakeResponse(b"{not-json"),
        )
        with self.assertRaisesRegex(
            RuntimeError, "OpenAI image generation returned invalid JSON"
        ):
            bad_json_provider.generate("fox")

        empty_data_provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: _FakeResponse(b'{"data": []}'),
        )
        with self.assertRaisesRegex(
            RuntimeError, "OpenAI image generation returned no image data"
        ):
            empty_data_provider.generate("fox")

        wrong_shape_provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: _FakeResponse(b'{"data": [1]}'),
        )
        with self.assertRaisesRegex(
            RuntimeError, "OpenAI image generation returned an unexpected response"
        ):
            wrong_shape_provider.generate("fox")

        invalid_b64_provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: _FakeResponse(
                b'{"data": [{"b64_json": "a"}]}'
            ),
        )
        with self.assertRaisesRegex(
            RuntimeError, "OpenAI image generation returned invalid base64 data"
        ):
            invalid_b64_provider.generate("fox")

        missing_payload_provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: _FakeResponse(
                b'{"data": [{"revised_prompt": "fox"}]}'
            ),
        )
        with self.assertRaisesRegex(
            RuntimeError, "OpenAI image generation returned no usable image payload"
        ):
            missing_payload_provider.generate("fox")

    def test_openai_image_provider_download_reports_errors(self) -> None:
        provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: (_ for _ in ()).throw(URLError("offline")),
        )
        with self.assertRaisesRegex(
            RuntimeError, "OpenAI image download failed: offline"
        ):
            provider._download_response_bytes("https://example.com/x.png")

        http_error = HTTPError(
            url="https://example.com/x.png",
            code=404,
            msg="Not Found",
            hdrs=None,
            fp=None,
        )
        http_error.read = lambda: b""  # type: ignore[assignment]
        provider = OpenAIImageProvider(
            api_key="test-key",
            urlopen=lambda request, timeout: (_ for _ in ()).throw(http_error),
        )
        with self.assertRaisesRegex(RuntimeError, "OpenAI image download failed: 404"):
            provider._download_response_bytes("https://example.com/x.png")

    def test_media_resolver_downloads_and_caches_remote_images(self) -> None:
        calls: list[tuple[object, float]] = []

        def fake_urlopen(request: object, timeout: float) -> _FakeResponse:
            calls.append((request, timeout))
            return _FakeResponse(b"image-data", content_type="image/png")

        with tempfile.TemporaryDirectory() as tmpdir:
            resolver = MediaResolver(cache_dir=tmpdir, urlopen=fake_urlopen)
            resolved = resolver.resolve_image(
                {"url": "https://example.com/assets/plot.png", "caption": "Remote"}
            )
            cached_path = resolved["path"]
            self.assertIsInstance(cached_path, Path)
            self.assertEqual(cached_path.read_bytes(), b"image-data")
            self.assertEqual(resolved["caption"], "Remote")

            resolved_again = resolver.resolve_image(
                {"url": "https://example.com/assets/plot.png"}
            )
            self.assertEqual(resolved_again["path"], cached_path)
            self.assertEqual(len(calls), 1)

    def test_media_resolver_download_handles_header_mapping_and_cache_extension(
        self,
    ) -> None:
        calls: list[tuple[object, float]] = []
        url = "https://example.com/render"

        def fake_urlopen(request: object, timeout: float) -> _FakeResponse:
            calls.append((request, timeout))
            response = _FakeResponse(b"image-data")
            response.headers = _FakeHeadersMapping({"Content-Type": "image/png"})
            return response

        with tempfile.TemporaryDirectory() as tmpdir:
            resolver = MediaResolver(cache_dir=tmpdir, urlopen=fake_urlopen)
            digest = hashlib.sha256(f"url:{url}".encode("utf-8")).hexdigest()
            cached_png = Path(tmpdir) / f"url-{digest}.png"
            cached_png.write_bytes(b"cached-image")

            resolved = resolver.resolve_image({"url": url})
            self.assertEqual(resolved["path"], cached_png)
            self.assertEqual(cached_png.read_bytes(), b"cached-image")
            self.assertEqual(len(calls), 1)

    def test_media_resolver_download_reports_errors(self) -> None:
        resolver = MediaResolver(
            urlopen=lambda request, timeout: (_ for _ in ()).throw(URLError("offline"))
        )
        with self.assertRaisesRegex(RuntimeError, "image download failed: offline"):
            resolver.resolve_image({"url": "https://example.com/a.png"})

        http_error = HTTPError(
            url="https://example.com/a.png",
            code=500,
            msg="Internal Server Error",
            hdrs=None,
            fp=None,
        )
        http_error.read = lambda: b""  # type: ignore[assignment]
        resolver = MediaResolver(
            urlopen=lambda request, timeout: (_ for _ in ()).throw(http_error)
        )
        with self.assertRaisesRegex(RuntimeError, "image download failed: 500"):
            resolver.resolve_image({"url": "https://example.com/a.png"})

    def test_media_resolver_generates_and_caches_prompt_images(self) -> None:
        provider = _FakePromptProvider()

        with tempfile.TemporaryDirectory() as tmpdir:
            resolver = MediaResolver(image_provider=provider, cache_dir=tmpdir)
            resolved = resolver.resolve_image(
                {
                    "prompt": "diagram of a data pipeline",
                    "size": "1024x1024",
                    "output_format": "png",
                    "caption": "Generated",
                }
            )

            cached_path = resolved["path"]
            self.assertIsInstance(cached_path, Path)
            self.assertEqual(cached_path.read_bytes(), b"generated-image")
            self.assertEqual(provider.calls[0][0], "diagram of a data pipeline")
            self.assertEqual(provider.calls[0][1]["size"], "1024x1024")
            self.assertEqual(resolved["caption"], "Generated")

            resolved_again = resolver.resolve_image(
                {
                    "prompt": "diagram of a data pipeline",
                    "size": "1024x1024",
                    "output_format": "png",
                }
            )
            self.assertEqual(resolved_again["path"], cached_path)
            self.assertEqual(len(provider.calls), 1)

    def test_media_resolver_handles_raw_values_and_validation_types(self) -> None:
        resolver = MediaResolver(image_provider=_FakePromptProvider())

        self.assertEqual(resolver.resolve_image("local.png"), "local.png")
        self.assertEqual(resolver.resolve_image(Path("local.png")), Path("local.png"))
        self.assertEqual(resolver.resolve_image(123), 123)

        with self.assertRaisesRegex(TypeError, "image url must be a string"):
            resolver.resolve_image({"url": 1})  # type: ignore[arg-type]

        with self.assertRaisesRegex(TypeError, "image prompt must be a string"):
            resolver.resolve_image({"prompt": 1})  # type: ignore[arg-type]

        with self.assertRaisesRegex(TypeError, "image quality must be a string"):
            resolver.resolve_image({"prompt": "diagram", "quality": 1})

    def test_media_resolver_validates_image_sources_and_prompt_requirements(
        self,
    ) -> None:
        resolver = MediaResolver()

        with self.assertRaisesRegex(
            ValueError,
            "exactly one of path, src, url, or prompt",
        ):
            resolver.resolve_image(
                {"path": "a.png", "url": "https://example.com/a.png"}
            )

        with self.assertRaisesRegex(
            ValueError,
            "exactly one of path, src, url, or prompt",
        ):
            resolver.resolve_image({"caption": "missing source"})

        with self.assertRaisesRegex(
            ValueError,
            "image prompt requires an image_provider",
        ):
            resolver.resolve_image({"prompt": "mountain landscape"})

        with self.assertRaisesRegex(
            TypeError,
            "image output_compression must be an integer",
        ):
            MediaResolver(image_provider=_FakePromptProvider()).resolve_image(
                {"prompt": "diagram", "output_compression": "80"}
            )


def test_openai_image_provider_uses_env_key(openai_api_key: str | None) -> None:
    if not openai_api_key:
        pytest.skip("OPENAI_API_KEY is not configured in tests/.env")

    provider = OpenAIImageProvider(
        urlopen=lambda request, timeout: _FakeResponse(b'{"data": []}')
    )
    assert provider.cache_key == "openai:https://api.openai.com/v1:gpt-image-1"


if __name__ == "__main__":
    unittest.main()
