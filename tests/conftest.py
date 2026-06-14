"""Shared pytest fixtures for the EntraID MCP server resource tests.

These tests exercise the resource functions against a fully mocked Microsoft
Graph client. No live Graph calls are made.
"""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from msgraph.generated.models.group import Group
from msgraph.generated.models.o_data_errors.o_data_error import ODataError


def make_odata_error(status_code: int, message: str = "") -> ODataError:
    """Build an ODataError (the SDK error type, a subclass of kiota APIError)
    with a concrete HTTP status code and message, mirroring what the Graph
    request adapter raises on a failed request."""
    err = ODataError()
    err.response_status_code = status_code
    err.message = message
    return err


class FakeGraphClient:
    """Stands in for utils.graph_client.GraphClient.

    The resource functions only call ``graph_client.get_client()`` to obtain the
    underlying GraphServiceClient, so we expose exactly that.
    """

    def __init__(self, service_client: MagicMock):
        self._service_client = service_client

    def get_client(self) -> MagicMock:
        return self._service_client


@pytest.fixture
def graph_setup():
    """Provide a mocked Graph service client with addressable leaf coroutines.

    Returns a dict exposing the fluent-chain leaf mocks so individual tests can
    assert on / configure them:

      - graph_client:        the FakeGraphClient passed to resource functions
      - service_client:      the underlying mocked GraphServiceClient
      - group_get:           AsyncMock for groups.by_group_id(id).get()  (dynamic-group guard)
      - member_ref_delete:   AsyncMock for ...members.by_directory_object_id(id).ref.delete()
      - member_ref_post:     AsyncMock for ...members.ref.post()
      - owner_ref_post:      AsyncMock for ...owners.ref.post()
      - owner_ref_delete:    AsyncMock for ...owners.by_directory_object_id(id).ref.delete()
      - member_by_id_get:    AsyncMock for the (unsupported) single-member GET --
                             present so we can prove it is NOT called.
    """
    service_client = MagicMock(name="GraphServiceClient")

    # groups.by_group_id(group_id) -> group_item builder
    group_item = MagicMock(name="group_item")
    service_client.groups.by_group_id.return_value = group_item

    # Dynamic-group guard: groups.by_group_id(id).get()
    group_get = AsyncMock(name="group_get", return_value=Group())
    group_item.get = group_get

    # --- members navigation ---
    members = group_item.members

    # members.by_directory_object_id(member_id) -> member_item builder
    member_item = MagicMock(name="member_item")
    members.by_directory_object_id.return_value = member_item

    # The UNSUPPORTED single-member GET. Default: always raises (real Graph
    # behaviour for GET /groups/{id}/members/{member-id}). Tests assert it is
    # never invoked.
    member_by_id_get = AsyncMock(
        name="member_by_id_get",
        side_effect=make_odata_error(400, "single-member GET is not a supported route"),
    )
    member_item.get = member_by_id_get

    # members.by_directory_object_id(member_id).ref.delete()
    member_ref_delete = AsyncMock(name="member_ref_delete", return_value=None)
    member_item.ref.delete = member_ref_delete

    # members.ref.post()
    member_ref_post = AsyncMock(name="member_ref_post", return_value=None)
    members.ref.post = member_ref_post

    # --- owners navigation ---
    owners = group_item.owners

    owner_ref_post = AsyncMock(name="owner_ref_post", return_value=None)
    owners.ref.post = owner_ref_post

    owner_item = MagicMock(name="owner_item")
    owners.by_directory_object_id.return_value = owner_item
    owner_ref_delete = AsyncMock(name="owner_ref_delete", return_value=None)
    owner_item.ref.delete = owner_ref_delete

    graph_client = FakeGraphClient(service_client)

    return {
        "graph_client": graph_client,
        "service_client": service_client,
        "group_item": group_item,
        "group_get": group_get,
        "member_ref_delete": member_ref_delete,
        "member_ref_post": member_ref_post,
        "member_by_id_get": member_by_id_get,
        "owner_ref_post": owner_ref_post,
        "owner_ref_delete": owner_ref_delete,
    }
