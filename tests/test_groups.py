"""Regression tests for group membership/ownership writes (DEV-2460).

Root cause covered here: ``remove_group_member`` used to do a pre-check
``GET /groups/{id}/members/{member-id}`` -- an unsupported Graph route that
ALWAYS raises -- and its ``except`` returned ``True``, so the real
``$ref`` delete below was never reached. The tool reported success while the
member stayed in the group.

These tests prove (a) the ``$ref`` delete IS invoked (no skipping), (b) the
broken single-member pre-check GET is NOT invoked, (c) a 404 is treated as
idempotent success, and (d) any other error propagates as a real failure.
"""

from __future__ import annotations

import pytest

from msgraph_mcp_server.resources import groups
from msgraph.generated.models.group import Group
from msgraph.generated.models.reference_create import ReferenceCreate

from .conftest import make_odata_error

GROUP_ID = "11111111-1111-1111-1111-111111111111"
MEMBER_ID = "22222222-2222-2222-2222-222222222222"
OWNER_ID = "33333333-3333-3333-3333-333333333333"


# ---------------------------------------------------------------------------
# remove_group_member -- the DEV-2460 regression
# ---------------------------------------------------------------------------


async def test_remove_group_member_calls_ref_delete(graph_setup):
    """The $ref delete MUST be invoked (regression: it used to be skipped)."""
    gc = graph_setup["graph_client"]

    result = await groups.remove_group_member(gc, GROUP_ID, MEMBER_ID)

    assert result is True
    # The actual write happened exactly once.
    graph_setup["member_ref_delete"].assert_awaited_once()
    # And the broken single-member pre-check GET was NEVER used.
    graph_setup["member_by_id_get"].assert_not_called()


async def test_remove_group_member_addresses_correct_ids(graph_setup):
    """The delete targets the requested group + member."""
    gc = graph_setup["graph_client"]

    await groups.remove_group_member(gc, GROUP_ID, MEMBER_ID)

    sc = graph_setup["service_client"]
    sc.groups.by_group_id.assert_any_call(GROUP_ID)
    sc.groups.by_group_id.return_value.members.by_directory_object_id.assert_called_once_with(MEMBER_ID)


async def test_remove_group_member_404_is_idempotent_success(graph_setup):
    """A 404 (member already not in the group) is idempotent success."""
    gc = graph_setup["graph_client"]
    graph_setup["member_ref_delete"].side_effect = make_odata_error(404, "not found")

    result = await groups.remove_group_member(gc, GROUP_ID, MEMBER_ID)

    assert result is True
    graph_setup["member_ref_delete"].assert_awaited_once()


async def test_remove_group_member_other_error_raises(graph_setup):
    """A non-404 error is a REAL failure and must propagate (no false success)."""
    gc = graph_setup["graph_client"]
    graph_setup["member_ref_delete"].side_effect = make_odata_error(403, "forbidden")

    with pytest.raises(Exception):
        await groups.remove_group_member(gc, GROUP_ID, MEMBER_ID)


async def test_remove_group_member_dynamic_group_rejected(graph_setup):
    """The dynamic-membership guard is preserved (ValueError, no delete)."""
    gc = graph_setup["graph_client"]
    dynamic = Group()
    dynamic.group_types = ["DynamicMembership"]
    graph_setup["group_get"].return_value = dynamic

    with pytest.raises(ValueError):
        await groups.remove_group_member(gc, GROUP_ID, MEMBER_ID)

    graph_setup["member_ref_delete"].assert_not_called()


# ---------------------------------------------------------------------------
# add_group_member -- same anti-pattern (dead pre-check GET) was removed
# ---------------------------------------------------------------------------


async def test_add_group_member_calls_ref_post(graph_setup):
    gc = graph_setup["graph_client"]

    result = await groups.add_group_member(gc, GROUP_ID, MEMBER_ID)

    assert result is True
    graph_setup["member_ref_post"].assert_awaited_once()
    # No dead single-member pre-check GET.
    graph_setup["member_by_id_get"].assert_not_called()


async def test_add_group_member_uses_reference_create(graph_setup):
    """The post body is a ReferenceCreate with the directoryObjects @odata.id."""
    gc = graph_setup["graph_client"]

    await groups.add_group_member(gc, GROUP_ID, MEMBER_ID)

    args, _ = graph_setup["member_ref_post"].await_args
    body = args[0]
    assert isinstance(body, ReferenceCreate)
    assert body.odata_id.endswith(f"/directoryObjects/{MEMBER_ID}")


async def test_add_group_member_already_member_is_idempotent(graph_setup):
    """Adding an existing member (HTTP 400 'already exist') is idempotent success."""
    gc = graph_setup["graph_client"]
    graph_setup["member_ref_post"].side_effect = make_odata_error(
        400, "One or more added object references already exist for the following modified properties: 'members'."
    )

    result = await groups.add_group_member(gc, GROUP_ID, MEMBER_ID)

    assert result is True


async def test_add_group_member_other_error_raises(graph_setup):
    gc = graph_setup["graph_client"]
    graph_setup["member_ref_post"].side_effect = make_odata_error(403, "forbidden")

    with pytest.raises(Exception):
        await groups.add_group_member(gc, GROUP_ID, MEMBER_ID)


async def test_add_group_member_dynamic_group_rejected(graph_setup):
    gc = graph_setup["graph_client"]
    dynamic = Group()
    dynamic.group_types = ["DynamicMembership"]
    graph_setup["group_get"].return_value = dynamic

    with pytest.raises(ValueError):
        await groups.add_group_member(gc, GROUP_ID, MEMBER_ID)

    graph_setup["member_ref_post"].assert_not_called()


# ---------------------------------------------------------------------------
# owners
# ---------------------------------------------------------------------------


async def test_add_group_owner_calls_ref_post_with_reference_create(graph_setup):
    gc = graph_setup["graph_client"]

    result = await groups.add_group_owner(gc, GROUP_ID, OWNER_ID)

    assert result is True
    graph_setup["owner_ref_post"].assert_awaited_once()
    args, _ = graph_setup["owner_ref_post"].await_args
    body = args[0]
    assert isinstance(body, ReferenceCreate)
    assert body.odata_id.endswith(f"/directoryObjects/{OWNER_ID}")


async def test_add_group_owner_already_owner_is_idempotent(graph_setup):
    gc = graph_setup["graph_client"]
    graph_setup["owner_ref_post"].side_effect = make_odata_error(
        400, "added object references already exist"
    )

    result = await groups.add_group_owner(gc, GROUP_ID, OWNER_ID)

    assert result is True


async def test_add_group_owner_other_error_raises(graph_setup):
    gc = graph_setup["graph_client"]
    graph_setup["owner_ref_post"].side_effect = make_odata_error(403, "forbidden")

    with pytest.raises(Exception):
        await groups.add_group_owner(gc, GROUP_ID, OWNER_ID)


async def test_remove_group_owner_calls_ref_delete(graph_setup):
    gc = graph_setup["graph_client"]

    result = await groups.remove_group_owner(gc, GROUP_ID, OWNER_ID)

    assert result is True
    graph_setup["owner_ref_delete"].assert_awaited_once()


async def test_remove_group_owner_404_is_idempotent(graph_setup):
    gc = graph_setup["graph_client"]
    graph_setup["owner_ref_delete"].side_effect = make_odata_error(404, "not found")

    result = await groups.remove_group_owner(gc, GROUP_ID, OWNER_ID)

    assert result is True


async def test_remove_group_owner_other_error_raises(graph_setup):
    gc = graph_setup["graph_client"]
    graph_setup["owner_ref_delete"].side_effect = make_odata_error(500, "server error")

    with pytest.raises(Exception):
        await groups.remove_group_owner(gc, GROUP_ID, OWNER_ID)
