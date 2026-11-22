import test_support
import unittest
from spam_guard.group_calendar_guard import matches, in_scope

class GuardTests(unittest.TestCase):
    def event(self, address, own=False):
        return {'organizer': {'emailAddress': {'address':address}}, 'isOrganizer':own, 'start':{'dateTime':'2000-01-18T12:00:00', 'timeZone':'Australia/Melbourne'}}
    def test_exact_domain(self):
        self.assertTrue(matches(self.event('test@GROUPS.OUTLOOK.COM')))
        self.assertFalse(matches(self.event('test@groups.outlook.com.evil.example')))
    def test_self_protected(self):
        self.assertFalse(matches(self.event('test@groups.outlook.com', True)))
    def test_history_and_series_protected(self):
        event=self.event('test@groups.outlook.com')
        self.assertTrue(in_scope(event))
        event['type']='seriesMaster'
        self.assertFalse(in_scope(event))
        event.pop('type');event['start']['dateTime']='2000-01-16T12:00:00'
        self.assertFalse(in_scope(event))

class FakeClient:
    def __init__(self, contacts_fail=False):
        self.deleted=False;self.contacts_fail=contacts_fail
        self.event={'id':'synthetic','organizer':{'emailAddress':{'address':'unknown@example.com'}},'isOrganizer':False,'start':{'dateTime':'2000-01-18T12:00:00','timeZone':'Australia/Melbourne'}}
    def request(self, method, path):
        from spam_guard.group_calendar_guard import GuardError
        if path.startswith('/me/?'): return {'mail':'owner@example.com'}
        if method=='DELETE': self.deleted=True;return {}
        if self.deleted: raise GuardError('HTTP_404')
        return self.event
    def pages(self,path):
        from spam_guard.group_calendar_guard import GuardError
        if path.startswith('/me/contacts'):
            if self.contacts_fail: raise GuardError('HTTP_403')
            return iter([])
        if path.startswith('/me/calendars?'): return iter([{'id':'owned','owner':{'address':'owner@example.com'},'canEdit':True}])
        return iter([self.event])

class ScanTests(unittest.TestCase):
    def test_account_not_configured_fails_closed(self):
        from unittest.mock import patch
        from spam_guard import group_calendar_guard as guard
        client=FakeClient()
        with patch.object(guard.config, 'ACCOUNT', ''):
            with self.assertRaisesRegex(guard.GuardError, 'ACCOUNT_NOT_CONFIGURED'):
                guard.scan(client, apply=True)
        self.assertFalse(client.deleted)

    def test_unknown_live_deleted_and_verified(self):
        import tempfile
        from pathlib import Path
        from unittest.mock import patch
        from spam_guard import group_calendar_guard as guard
        client=FakeClient()
        with tempfile.TemporaryDirectory() as directory, patch.object(guard,'STATE',Path(directory)):
            self.assertEqual(guard.scan(client,apply=True,contacts_apply=True)['deleted'],1)
            self.assertTrue((Path(directory)/'deletions.jsonl').exists())
    def test_contacts_failure_no_delete(self):
        from spam_guard import group_calendar_guard as guard
        client=FakeClient(contacts_fail=True)
        with self.assertRaises(guard.GuardError): guard.scan(client,apply=True,contacts_apply=True)
        self.assertFalse(client.deleted)
    def test_self_preserved_under_unknown_policy(self):
        from spam_guard import group_calendar_guard as guard
        client=FakeClient();client.event['isOrganizer']=True
        self.assertEqual(guard.scan(client,apply=True,contacts_apply=True)['deleted'],0)

class RecurringClient(FakeClient):
    def __init__(self):
        super().__init__()
        self.event.update(type='seriesMaster', recurrence={'range':{'startDate':'2000-01-01'}})
        self.event['start']['dateTime']='2000-01-01T12:00:00'
        import copy
        self.items={}
        for key, day, own in [('past','2000-01-16',False),('future','2000-01-18',False),('self','2000-01-19',True)]:
            item=copy.deepcopy(self.event);item.update(id=key,type='occurrence',isOrganizer=own)
            item['start']['dateTime']=day+'T12:00:00';self.items[key]=item
        self.removed=[]
    def pages(self,path):
        if '/instances?' in path: return iter(self.items.values())
        return super().pages(path)
    def request(self,method,path):
        from spam_guard.group_calendar_guard import GuardError
        if path.startswith('/me/?'): return {'mail':'owner@example.com'}
        key=path.split('?')[0].rsplit('/',1)[-1]
        if method=='DELETE': self.removed.append(key);return {}
        if key in self.removed: raise GuardError('HTTP_404')
        return self.items.get(key,self.event)

class RecurrenceTests(unittest.TestCase):
    def test_new_master_can_be_removed_entirely(self):
        from spam_guard.group_calendar_guard import in_scope
        event=RecurringClient().event
        event['recurrence']['range']['startDate']='2000-01-17'
        event['start']['dateTime']='2000-01-17T12:00:00'
        self.assertTrue(in_scope(event))
    def test_missing_recurrence_never_deletes_master(self):
        from spam_guard.group_calendar_guard import in_scope
        event=RecurringClient().event;event.pop('recurrence')
        self.assertFalse(in_scope(event))
    def test_old_series_only_future_external_occurrence_deleted(self):
        from spam_guard import group_calendar_guard as guard
        from unittest.mock import patch
        from pathlib import Path
        import tempfile
        client=RecurringClient()
        with tempfile.TemporaryDirectory() as directory, patch.object(guard,'STATE',Path(directory)):
            result=guard.scan(client,apply=True,contacts_apply=True)
        self.assertEqual(client.removed,['future'])
        self.assertEqual(result['series_expanded'],1)
        self.assertEqual(result['instances_scanned'],3)
    def test_preview_expands_without_mutation(self):
        from spam_guard import group_calendar_guard as guard
        client=RecurringClient()
        result=guard.scan(client,contacts_preview=True)
        self.assertEqual(result['unknown_organizers'],1)
        self.assertEqual(client.removed,[])
    def test_new_master_deleted_as_whole(self):
        from spam_guard import group_calendar_guard as guard
        from unittest.mock import patch
        from pathlib import Path
        import tempfile
        client=FakeClient()
        client.event.update(type='seriesMaster',recurrence={'range':{'startDate':'2000-01-18'}})
        with tempfile.TemporaryDirectory() as directory, patch.object(guard,'STATE',Path(directory)):
            result=guard.scan(client,apply=True,contacts_apply=True)
        self.assertTrue(client.deleted)
        self.assertEqual(result['deleted'],1)
        self.assertEqual(result['series_expanded'],0)
