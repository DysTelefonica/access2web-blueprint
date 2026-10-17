#!/usr/bin/env python3
# ci-pattern asset tests — local GitHub API stub (DysTelefonica/team-skills#694, CodeQL py/http-response-splitting)
"""The stub's ``Link`` header carries only data the server already holds: its
own listening address and the SHA key stored in ``state['pulls_pages']``. Nothing
the client sends (``Host``, request path, SHA segment) is copied into a response
header. Raw sockets are used because ``http.client`` refuses to send a header
value with a line break."""
import http.client
import json
import socket
import tempfile
import unittest
import urllib.parse
from pathlib import Path

from _github_stub import github_stub

SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
PAGE_ONE, PAGE_TWO = [{"number": 1}], [{"number": 2}]
PATH = f"/repos/o/r/commits/{SHA}/pulls"


def paged_state() -> dict:
    return {"calls": [], "pulls": {}, "pulls_pages": {SHA: [PAGE_ONE, PAGE_TWO]}}


def raw_get(port: int, target: str, header_lines: list[str]) -> tuple[str, bytes]:
    request = f"GET {target} HTTP/1.1\r\n" + "".join(f"{line}\r\n" for line in header_lines) + "Connection: close\r\n\r\n"
    chunks = []
    with socket.create_connection(("127.0.0.1", port), timeout=5) as conn:
        conn.sendall(request.encode("latin-1"))
        while chunk := conn.recv(4096):
            chunks.append(chunk)
    head, _, body = b"".join(chunks).partition(b"\r\n\r\n")
    return head.decode("latin-1"), body


class TestLinkCarriesOnlyServerSideData(unittest.TestCase):
    def serve(self, target: str, header_lines: list[str]) -> tuple[str, bytes, str]:
        with tempfile.TemporaryDirectory() as tmp, github_stub(Path(tmp, "state.json"), paged_state()) as url:
            head, body = raw_get(int(url.rsplit(":", 1)[1]), target, header_lines)
            return head, body, url

    def link_of(self, head: str) -> str:
        links = [line for line in head.split("\r\n") if line.lower().startswith("link:")]
        self.assertEqual(len(links), 1, head)
        return links[0]

    def expected_link(self, url: str) -> str:
        return f'Link: <{url}/repos/stub/stub/commits/{SHA}/pulls?page=2>; rel="next"'

    def test_link_uses_the_server_address_and_the_stored_sha(self) -> None:
        head, _, url = self.serve(PATH, ["Host: 127.0.0.1"])
        self.assertEqual(self.link_of(head), self.expected_link(url))

    def test_a_foreign_host_header_is_not_reflected(self) -> None:
        head, _, url = self.serve(PATH, ["Host: evil.example"])
        self.assertNotIn("evil.example", head)
        self.assertEqual(self.link_of(head), self.expected_link(url))

    def test_a_folded_host_header_cannot_inject_a_response_header(self) -> None:
        head, _, url = self.serve(PATH, ["Host: evil.example", " X-Injected: yes"])
        self.assertFalse([line for line in head.split("\r\n") if line.lower().startswith("x-injected")], head)
        self.assertEqual(self.link_of(head), self.expected_link(url))

    def test_a_request_without_host_still_gets_the_servers_own_address(self) -> None:
        head, _, url = self.serve(PATH, [])
        self.assertEqual(self.link_of(head), self.expected_link(url))

    def test_odd_characters_in_the_request_path_are_not_reflected(self) -> None:
        target = f"/repos/o%0d%0aX-Injected:%20yes/r%20%3Cx%3E;a=b/commits/{SHA}/pulls"
        head, _, url = self.serve(target, ["Host: 127.0.0.1"])
        self.assertNotIn("X-Injected", head)
        self.assertNotIn("%0d", head.lower())
        self.assertNotIn("%3c", head.lower())
        self.assertEqual(self.link_of(head), self.expected_link(url))

    def test_a_forged_sha_segment_is_not_reflected_and_gets_no_link(self) -> None:
        head, body, _ = self.serve(f"/repos/o/r/commits/{SHA}%0d%0aX-Injected:%20yes/pulls", ["Host: 127.0.0.1"])
        self.assertNotIn("X-Injected", head)
        self.assertNotIn("link:", head.lower())
        self.assertEqual(json.loads(body), [])

    def test_following_the_link_still_pages_through_the_results(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, github_stub(Path(tmp, "state.json"), paged_state()) as url:
            port = int(url.rsplit(":", 1)[1])
            head, body = raw_get(port, PATH, ["Host: 127.0.0.1"])
            self.assertEqual(json.loads(body), PAGE_ONE)
            next_url = self.link_of(head).split("<", 1)[1].split(">", 1)[0]
            parts = urllib.parse.urlsplit(next_url)
            conn = http.client.HTTPConnection(parts.hostname, parts.port, timeout=5)
            conn.request("GET", f"{parts.path}?{parts.query}")
            second = conn.getresponse()
            self.assertEqual(json.loads(second.read()), PAGE_TWO)
            self.assertIsNone(second.getheader("Link"))
            conn.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
