import test_support
import unittest
from unittest.mock import AsyncMock, patch
from spam_guard import leave_groups_via_cdp as groups

class GroupTests(unittest.IsolatedAsyncioTestCase):
    async def test_identity_check_uses_local_account_configuration(self):
        with patch.object(groups.config, 'ACCOUNT', 'owner@example.com'):
            self.assertTrue(groups.identity_is_expected('owner@example.com'))
            self.assertFalse(groups.identity_is_expected('other@example.com'))
        with patch.object(groups.config, 'ACCOUNT', ''):
            self.assertIsNone(groups.identity_is_expected('owner@example.com'))

    async def test_missing_button_is_nonfatal_outcome(self):
        with patch.object(groups,'cdp_send',new_callable=AsyncMock), patch.object(groups,'_wait_for_ready',new_callable=AsyncMock), patch.object(groups,'page_text',new_callable=AsyncMock,return_value=''), patch.object(groups.asyncio,'sleep',new_callable=AsyncMock), patch.object(groups,'evaluate',new_callable=AsyncMock,side_effect=[True,True,False]):
            result=await groups.leave_one(None,{'gid':'synthetic','name':'Synthetic'},120)
        self.assertEqual(result,'leave_unavailable')
    async def test_final_group_empty_tree_verified_after_reload(self):
        with patch.object(groups,'cdp_send',new_callable=AsyncMock), patch.object(groups,'_wait_for_ready',new_callable=AsyncMock), patch.object(groups,'page_text',new_callable=AsyncMock,return_value=''), patch.object(groups.asyncio,'sleep',new_callable=AsyncMock), patch.object(groups,'evaluate',new_callable=AsyncMock,side_effect=[True,True,True,True]), patch.object(groups,'_discover',new_callable=AsyncMock,return_value=[]) as discover:
            result=await groups.leave_one(None,{'gid':'synthetic','name':'Synthetic'},120)
        self.assertEqual(result,'left')
        self.assertEqual(discover.await_count,2)
    async def test_no_named_exceptions(self):
        result=groups.summarize([{'gid':'synthetic','name':'Synthetic'},{'gid':'other','name':'Other'}])
        self.assertEqual(result,{'discovered':2,'targets':2})
    async def test_unavailable_does_not_alert(self):
        self.assertEqual(groups.failure_code({'leave_unavailable':1,'errors':0}), '')
    async def test_errors_not_success(self):
        self.assertEqual(groups.failure_code({'errors':1}),'GROUP_ERRORS')
    async def test_group_error_alert_is_configurable(self):
        self.assertTrue(groups.should_notify('GROUP_ERRORS',{'notify_group_errors':True}))
        self.assertFalse(groups.should_notify('GROUP_ERRORS',{'notify_group_errors':False}))
        self.assertTrue(groups.should_notify('LOGIN_REQUIRED',{'notify_group_errors':False}))
        self.assertFalse(groups.should_notify('',{'notify_group_errors':True}))
    async def test_config_defaults_to_alerting(self):
        import tempfile
        from pathlib import Path
        with tempfile.TemporaryDirectory() as tmp, patch.object(groups,'STATE_DIR',Path(tmp)):
            self.assertTrue(groups.load_config()['notify_group_errors'])
            (Path(tmp)/'config.json').write_text('not json')
            self.assertTrue(groups.load_config()['notify_group_errors'])
            (Path(tmp)/'config.json').write_text('{"notify_group_errors": false}')
            self.assertFalse(groups.load_config()['notify_group_errors'])
    async def test_missing_leave_does_not_stop_other_groups(self):
        from unittest.mock import MagicMock
        ws=MagicMock();ws.close=AsyncMock()
        entries=[{'gid':'first','name':'First'},{'gid':'second','name':'Second'}]
        with patch.object(groups,'http_get_json',side_effect=[{},[{'type':'page','url':'https://outlook.live.com/groups/home','webSocketDebuggerUrl':'ws://127.0.0.1:9222/synthetic'}]]), patch.object(groups.websockets,'connect',new_callable=AsyncMock,return_value=ws), patch.object(groups,'cdp_send',new_callable=AsyncMock), patch.object(groups,'_wait_for_ready',new_callable=AsyncMock), patch.object(groups,'page_text',new_callable=AsyncMock,return_value=''), patch.object(groups,'verify_identity',new_callable=AsyncMock,return_value=None), patch.object(groups,'_discover',new_callable=AsyncMock,side_effect=[entries,[]]), patch.object(groups,'leave_one',new_callable=AsyncMock,side_effect=['leave_unavailable','left']) as leave, patch.object(groups.asyncio,'sleep',new_callable=AsyncMock):
            result=await groups.run(confirm=True)
        self.assertEqual(leave.await_count,2)
        self.assertEqual(result['leave_unavailable'],1)
        self.assertEqual(result['left'],1)
        self.assertEqual(result['errors'],0)
