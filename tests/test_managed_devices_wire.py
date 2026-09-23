"""Wire-level guards: a real GraphServiceClient whose HTTP transport is an
httpx MockTransport, so these assert on the exact request the SDK puts on the
wire (method, path, query, body) rather than on which mock was called.

- managed devices by user used `$filter=userId eq ...` on
  deviceManagement/managedDevices. Intune rejects that with 400 "Unsupported
  parameter", and it leaves `userId` empty on devices anyway (VOTACALL-ROY,
  measured live 2026-09-23), so even a working filter would return nothing.
- members/owners `$ref` POSTs: test_groups.py proves a ReferenceCreate is
  passed; this proves it serializes to the `@odata.id` body Graph requires
  (Graph rejected the old `{"id": ...}` body with 400 "EndOfInput").
"""

from __future__ import annotations

import asyncio
import json

import httpx
import pytest
from kiota_abstractions.authentication import AnonymousAuthenticationProvider
from msgraph import GraphRequestAdapter, GraphServiceClient

from msgraph_mcp_server.resources import groups, managed_devices

GROUP = "6ea957ba-2018-48a6-93ff-99623cf40250"
MEMBER = "8a93427a-8e4b-4d5b-bbc3-937ae3a08462"
USER = "aced28a0-8574-4cc9-8946-b910f38ec947"


class _StubGraphClient:
    def __init__(self, client):
        self._client = client

    def get_client(self):
        return self._client


@pytest.fixture
def wire():
    """Yields (graph_client, requests, route). Unrouted requests get a 404."""
    requests: list[httpx.Request] = []
    routes: dict[tuple[str, str], tuple[int, object]] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        for (method, suffix), (status, body) in routes.items():
            if request.method == method and request.url.path.endswith(suffix):
                return httpx.Response(status, json=body) if body is not None else httpx.Response(status)
        return httpx.Response(404, json={"error": {"code": "Request_ResourceNotFound", "message": "nf"}})

    adapter = GraphRequestAdapter(
        AnonymousAuthenticationProvider(),
        client=httpx.AsyncClient(transport=httpx.MockTransport(handler)),
    )

    def route(method, suffix, status=200, body=None):
        routes[(method, suffix)] = (status, body)

    yield _StubGraphClient(GraphServiceClient(request_adapter=adapter)), requests, route


def _body_json(request: httpx.Request):
    assert request.content, f"{request.method} {request.url.path} was sent with an EMPTY body"
    return json.loads(request.content)


def test_managed_devices_by_user_uses_user_navigation_not_filter(wire):
    client, requests, route = wire
    route("GET", f"/users/{USER}/managedDevices", body={"value": [
        {"id": "91e3cd0c-9053-4b95-97c9-9cb81b69f3be", "deviceName": "VOTACALL-ROY",
         "operatingSystem": "Windows", "userId": ""},
    ]})

    devices = asyncio.run(managed_devices.get_managed_devices_by_user(client, USER))

    # positive control: the parse returned the known device, so an empty list
    # can never pass as "user has no devices"
    assert [d["deviceName"] for d in devices] == ["VOTACALL-ROY"]
    gets = [r for r in requests if r.method == "GET"]
    assert len(gets) == 1
    assert gets[0].url.path.endswith(f"/users/{USER}/managedDevices"), gets[0].url
    assert "filter" not in str(gets[0].url).lower(), f"Intune rejects $filter here: {gets[0].url}"


@pytest.mark.parametrize("fn_name,collection", [
    ("add_group_member", "members"),
    ("add_group_owner", "owners"),
])
def test_ref_post_serializes_odata_id(wire, fn_name, collection):
    client, requests, route = wire
    route("GET", f"/groups/{GROUP}", body={"id": GROUP, "groupTypes": []})
    route("POST", f"/groups/{GROUP}/{collection}/$ref", status=204)

    assert asyncio.run(getattr(groups, fn_name)(client, GROUP, MEMBER)) is True

    posts = [r for r in requests if r.method == "POST"]
    # positive control: the write actually happened and hit the $ref endpoint
    assert len(posts) == 1, f"expected exactly one POST, saw {[(r.method, r.url.path) for r in requests]}"
    assert posts[0].url.path.endswith(f"/groups/{GROUP}/{collection}/$ref")
    assert _body_json(posts[0]).get("@odata.id") == f"https://graph.microsoft.com/v1.0/directoryObjects/{MEMBER}"
