import sys
import unittest
from pathlib import Path
from unittest.mock import patch
from pydantic import ValidationError

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'Server'))
import revit_mcp_server as server

class AnnotationReviewTests(unittest.TestCase):
    def test_schedule_clone_preserves_source_scope_and_rollback(self):
        item = server.RoomScheduleClone(source_schedule_unique_id='source', view_unique_id='view',
            link_instance_unique_id='link', name='Fragment', number_prefix='2.-1.')
        with patch.object(server.client, 'command', return_value={}) as send:
            server.clone_revit_room_schedule('doc', 3, [item], session_id='session')
            self.assertEqual(send.call_args.args, ('clone_room_schedule', 'session'))
            self.assertTrue(send.call_args.kwargs['dry_run'])
            self.assertEqual(send.call_args.kwargs['items'][0]['number_prefix'], '2.-1.')
        with self.assertRaises(ValidationError):
            server.RoomScheduleClone(**dict(item.model_dump(), number_prefix='2'))

    def test_delete_requires_view_and_defaults_to_rollback(self):
        with self.assertRaises(ValidationError):
            server.TagDelete(tag_unique_id='tag')
        item = server.TagDelete(view_unique_id='view', tag_unique_id='tag')
        with patch.object(server.client, 'command', return_value={}) as send:
            server.delete_revit_tags('doc', 3, [item], session_id='session')
            self.assertEqual(send.call_args.args, ('delete_tags', 'session'))
            self.assertTrue(send.call_args.kwargs['dry_run'])
            self.assertEqual(send.call_args.kwargs['expected_revision'], 3)
            self.assertEqual(send.call_args.kwargs['items'], [{'view_unique_id':'view', 'tag_unique_id':'tag'}])

    def test_pipe_endpoint_cannot_disable_leader(self):
        values = dict(view_unique_id='view', tag_unique_id='tag', expected_head_mm=[0,0,0], head_mm=[1,1,0], leader_end_mm=[2,2,0])
        with self.assertRaises(ValidationError):
            server.TagUpdate(**values, leader=False)
        with self.assertRaises(ValidationError):
            server.TagUpdate(**dict(values, leader_end_mm=[float('nan'),0,0]))
        with patch.object(server.client, 'command', return_value={}) as send:
            server.update_revit_tags('doc', 5, [server.TagUpdate(**values)], dry_run=False)
            self.assertEqual(send.call_args.kwargs['items'][0]['leader_end_mm'], [2,2,0])

    def test_space_context_is_read_only_and_keeps_explicit_radiators(self):
        with patch.object(server.client, 'command', return_value={}) as send:
            server.get_revit_space_annotation_context('doc', 'view', ['radiator'], session_id='session')
            send.assert_called_once_with('get_space_annotation_context', 'session', document_id='doc', view_unique_id='view', radiator_unique_ids=['radiator'])
