#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import json
import traceback

MODULE_PATH = r"C:\ProgramData\Blackmagic Design\DaVinci Resolve\Support\Developer\Scripting\Modules"
if MODULE_PATH not in sys.path:
    sys.path.append(MODULE_PATH)

import DaVinciResolveScript as dvr

try:
    import BlackmagicFusion as bmd
except Exception:
    bmd = None

COLOR_LIST = [
    "Orange", "Apricot", "Yellow", "Lime", "Olive", "Green",
    "Teal", "Navy", "Blue", "Purple", "Violet", "Pink",
    "Tan", "Beige", "Brown", "Chocolate"
]

DEFAULT_COLOR = "Orange"

DEFAULT_PRESERVE_OPTIONS = {
    "text": True,
    "center": True,
    "size": True,
    "angle": False,
    "pivot": False,
}


def get_settings_file_path():
    try:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "TextPlusStyleCopier_settings.json")
    except Exception:
        return "TextPlusStyleCopier_settings.json"


def load_ui_settings():
    settings_path = get_settings_file_path()
    default_settings = {
        "last_color": DEFAULT_COLOR,
        "keep_color": True,
        "target_track": 0,
        "preserve_options": dict(DEFAULT_PRESERVE_OPTIONS),
    }

    if not os.path.exists(settings_path):
        return default_settings

    try:
        with open(settings_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if data.get("last_color") not in COLOR_LIST:
            data["last_color"] = DEFAULT_COLOR

        if "keep_color" not in data:
            data["keep_color"] = True

        try:
            data["target_track"] = int(data.get("target_track", 0))
            if data["target_track"] < 0:
                data["target_track"] = 0
        except Exception:
            data["target_track"] = 0

        preserve_options = data.get("preserve_options", {})
        merged = dict(DEFAULT_PRESERVE_OPTIONS)
        for k in merged.keys():
            if k in preserve_options:
                merged[k] = bool(preserve_options[k])
        data["preserve_options"] = merged

        return data
    except Exception:
        return default_settings


def save_ui_settings(color, keep_color, preserve_options, target_track):
    settings_path = get_settings_file_path()

    merged = dict(DEFAULT_PRESERVE_OPTIONS)
    for k in merged.keys():
        if k in preserve_options:
            merged[k] = bool(preserve_options[k])

    try:
        target_track = int(target_track)
        if target_track < 0:
            target_track = 0
    except Exception:
        target_track = 0

    data = {
        "last_color": color if color in COLOR_LIST else DEFAULT_COLOR,
        "keep_color": bool(keep_color),
        "target_track": target_track,
        "preserve_options": merged,
    }

    try:
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Failed to save settings:", e)


def get_textplus_tool(fusion_comp):
    if fusion_comp is None:
        return None

    try:
        lst = fusion_comp.GetToolList(False, "TextPlus")
        if lst:
            for _, tool in lst.items():
                return tool
    except Exception:
        pass

    try:
        all_tools = fusion_comp.GetToolList(False) or {}
        for _, tool in all_tools.items():
            attrs = tool.GetAttrs() or {}
            if attrs.get("TOOLS_RegID") == "TextPlus":
                return tool
    except Exception:
        pass

    return None


def get_textplus_from_timeline_item(item):
    if not item:
        return None, None

    try:
        fusion_count = item.GetFusionCompCount()
    except Exception:
        fusion_count = 0

    if fusion_count == 0:
        return None, None

    try:
        comp = item.GetFusionCompByIndex(1)
    except Exception:
        return None, None

    tool = get_textplus_tool(comp)
    return comp, tool


def is_textplus_timeline_item(item):
    _, tool = get_textplus_from_timeline_item(item)
    return tool is not None


def get_selected_mediapool_clip(media_pool):
    if not media_pool:
        return None

    try:
        selected = media_pool.GetSelectedClips()
    except Exception:
        selected = None

    if not selected:
        return None

    if isinstance(selected, dict):
        for _, clip in selected.items():
            return clip
    elif isinstance(selected, list):
        if len(selected) > 0:
            return selected[0]

    return None


def get_timeline_video_items_in_track(timeline, track_index):
    try:
        return timeline.GetItemListInTrack("video", track_index) or []
    except Exception:
        return []


def get_all_timeline_video_items(timeline):
    result = []

    try:
        track_count = timeline.GetTrackCount("video")
    except Exception:
        return result

    for track_index in range(1, track_count + 1):
        result.extend(get_timeline_video_items_in_track(timeline, track_index))

    return result


def append_clip_to_timeline_and_detect_new_item(media_pool, timeline, media_pool_item):
    before_items = get_all_timeline_video_items(timeline)
    before_ids = set(id(x) for x in before_items)

    try:
        appended = media_pool.AppendToTimeline([media_pool_item])
    except Exception as e:
        print("AppendToTimeline failed:", e)
        return None

    if appended:
        if isinstance(appended, list) and len(appended) > 0:
            for obj in appended:
                if obj and is_textplus_timeline_item(obj):
                    return obj
            return appended[-1]

    after_items = get_all_timeline_video_items(timeline)
    new_items = [x for x in after_items if id(x) not in before_ids]

    if new_items:
        return new_items[-1]

    return None


def get_current_timecode_safe(timeline):
    try:
        if hasattr(timeline, "GetCurrentTimecode"):
            return timeline.GetCurrentTimecode()
    except Exception:
        pass
    return None


def set_current_timecode_safe(timeline, project, timecode):
    if not timecode:
        return False

    try:
        if hasattr(timeline, "SetCurrentTimecode"):
            ok = timeline.SetCurrentTimecode(timecode)
            if ok is None or ok is True:
                return True
    except Exception:
        pass

    try:
        if hasattr(project, "SetCurrentTimecode"):
            ok = project.SetCurrentTimecode(timecode)
            if ok is None or ok is True:
                return True
    except Exception:
        pass

    return False


def get_clip_color_safe(item):
    try:
        if hasattr(item, "GetClipColor"):
            return item.GetClipColor()
    except Exception:
        pass
    return None


def clear_clip_color_safe(item):
    try:
        if hasattr(item, "ClearClipColor"):
            item.ClearClipColor()
            return True
    except Exception:
        pass

    for value in ["", "None"]:
        try:
            item.SetClipColor(value)
            return True
        except Exception:
            pass

    return False


def get_input_safe(tool, name):
    try:
        return tool.GetInput(name)
    except Exception:
        return None


def set_input_safe(tool, name, value):
    try:
        tool.SetInput(name, value)
        return True
    except Exception:
        return False


def build_preserve_map(dst_tool, preserve_options):
    input_map = {
        "text": "StyledText",
        "center": "Center",
        "size": "Size",
        "angle": "Angle",
        "pivot": "Pivot",
    }

    preserved = {}

    for key, input_name in input_map.items():
        if preserve_options.get(key):
            preserved[input_name] = get_input_safe(dst_tool, input_name)

    return preserved


def restore_preserved_inputs(dst_tool, preserved):
    ok = True

    for input_name, value in preserved.items():
        if value is not None:
            if not set_input_safe(dst_tool, input_name, value):
                print(f"Failed to restore {input_name}")
                ok = False

    return ok


def copy_style_with_preserve_options(src_tool, dst_comp, dst_tool, preserve_options):
    try:
        src_settings = src_tool.SaveSettings()
    except Exception as e:
        print("SaveSettings failed:", e)
        return False

    preserved = build_preserve_map(dst_tool, preserve_options)

    try:
        dst_comp.Lock()
    except Exception:
        pass

    ok = True

    try:
        dst_tool.LoadSettings(src_settings)
    except Exception as e:
        print("LoadSettings failed:", e)
        ok = False

    if not restore_preserved_inputs(dst_tool, preserved):
        ok = False

    try:
        dst_comp.Unlock()
    except Exception:
        pass

    return ok


def delete_timeline_item_safe(timeline, item):
    if not item:
        return False

    try:
        if hasattr(timeline, "DeleteClips"):
            ok = timeline.DeleteClips([item], False)
            if ok is None or ok is True:
                return True
    except Exception:
        pass

    try:
        if hasattr(timeline, "DeleteClips"):
            ok = timeline.DeleteClips([item])
            if ok is None or ok is True:
                return True
    except Exception:
        pass

    return False


def find_colored_textplus_items(timeline, target_color, target_track=0, exclude_item=None):
    result = []

    try:
        track_count = timeline.GetTrackCount("video")
    except Exception:
        track_count = 0

    if track_count <= 0:
        return result

    if target_track and target_track > 0:
        track_indices = [target_track]
    else:
        track_indices = list(range(1, track_count + 1))

    for track_index in track_indices:
        for item in get_timeline_video_items_in_track(timeline, track_index):
            if exclude_item and item == exclude_item:
                continue

            color = get_clip_color_safe(item)
            if color != target_color:
                continue

            if is_textplus_timeline_item(item):
                result.append(item)

    return result


def run_apply(resolve, target_color, keep_color, preserve_options, target_track):
    result = {
        "ok": False,
        "message": "",
        "copied_count": 0,
        "failed_count": 0,
    }

    pm = resolve.GetProjectManager()
    project = pm.GetCurrentProject() if pm else None

    if not project:
        result["message"] = "No project is currently open."
        return result

    media_pool = project.GetMediaPool()
    timeline = project.GetCurrentTimeline()

    if not timeline:
        result["message"] = "No timeline is currently open."
        return result

    original_timecode = get_current_timecode_safe(timeline)

    src_media_pool_item = get_selected_mediapool_clip(media_pool)

    if not src_media_pool_item:
        result["message"] = "Please select one source Text+ clip in the Power Bin / Media Pool."
        return result

    ref_timeline_item = append_clip_to_timeline_and_detect_new_item(
        media_pool,
        timeline,
        src_media_pool_item
    )

    if not ref_timeline_item:
        result["message"] = "Failed to append the source clip to the timeline."
        return result

    try:
        set_current_timecode_safe(timeline, project, original_timecode)

        src_comp, src_tool = get_textplus_from_timeline_item(ref_timeline_item)

        if not src_tool:
            result["message"] = "No TextPlus tool was found in the source clip."
            return result

        target_items = find_colored_textplus_items(
            timeline,
            target_color,
            target_track=target_track,
            exclude_item=ref_timeline_item
        )

        if not target_items:
            if target_track > 0:
                result["message"] = f"No {target_color} Text+ clips found on V{target_track}."
            else:
                result["message"] = f"No {target_color} Text+ clips found."
            return result

        copied_count = 0
        failed_count = 0

        for item in target_items:
            dst_comp, dst_tool = get_textplus_from_timeline_item(item)

            if not dst_tool:
                failed_count += 1
                continue

            ok = copy_style_with_preserve_options(
                src_tool,
                dst_comp,
                dst_tool,
                preserve_options
            )

            if ok:
                copied_count += 1

                if not keep_color:
                    clear_clip_color_safe(item)
            else:
                failed_count += 1

        result["ok"] = True
        result["copied_count"] = copied_count
        result["failed_count"] = failed_count

        if target_track > 0:
            result["message"] = f"Done: V{target_track} Success {copied_count} / Failed {failed_count}"
        else:
            result["message"] = f"Done: Success {copied_count} / Failed {failed_count}"

        return result

    finally:
        delete_timeline_item_safe(timeline, ref_timeline_item)
        set_current_timecode_safe(timeline, project, original_timecode)


def show_persistent_ui(resolve):
    if bmd is None:
        print("BlackmagicFusion UIManager could not be loaded.")
        return

    saved_settings = load_ui_settings()

    default_color = saved_settings.get("last_color", DEFAULT_COLOR)
    default_keep_color = bool(saved_settings.get("keep_color", True))
    default_target_track = int(saved_settings.get("target_track", 0))
    default_preserve = saved_settings.get(
        "preserve_options",
        dict(DEFAULT_PRESERVE_OPTIONS)
    )

    try:
        fusion = resolve.Fusion()
        ui = fusion.UIManager
        dispatcher = bmd.UIDispatcher(ui)
    except Exception as e:
        print("UI initialization failed:", e)
        return

    window = dispatcher.AddWindow(
        {
            "ID": "TextPlusStyleWin",
            "WindowTitle": "TextPlusStyleCopier",
            "Geometry": [100, 100, 560, 500],
        },
        ui.VGroup([
            ui.Label({
                "Text": "Target Clip Color",
            }),

            ui.ComboBox({
                "ID": "colorCombo",
                "MinimumSize": [500, 34],
            }),

            ui.VGap(6),

            ui.Label({
                "Text": "Target Track Number (0 = All Tracks)",
            }),

            ui.LineEdit({
                "ID": "trackEdit",
                "Text": str(default_target_track),
                "MinimumSize": [160, 30],
            }),

            ui.VGap(6),

            ui.CheckBox({
                "ID": "keepColorCheck",
                "Text": "Keep Clip Color After Apply",
                "Checked": default_keep_color,
            }),

            ui.VGap(8),

            ui.Label({
                "Text": "Preserve These Properties",
            }),

            ui.HGroup([
                ui.CheckBox({
                    "ID": "preserveTextCheck",
                    "Text": "Text",
                    "Checked": bool(default_preserve.get("text", True)),
                }),
                ui.CheckBox({
                    "ID": "preserveCenterCheck",
                    "Text": "Position",
                    "Checked": bool(default_preserve.get("center", True)),
                }),
                ui.CheckBox({
                    "ID": "preserveSizeCheck",
                    "Text": "Size",
                    "Checked": bool(default_preserve.get("size", True)),
                }),
            ]),

            ui.HGroup([
                ui.CheckBox({
                    "ID": "preserveAngleCheck",
                    "Text": "Rotation",
                    "Checked": bool(default_preserve.get("angle", False)),
                }),
                ui.CheckBox({
                    "ID": "preservePivotCheck",
                    "Text": "Pivot",
                    "Checked": bool(default_preserve.get("pivot", False)),
                }),
            ]),

            ui.VGap(8),

            ui.Label({
                "ID": "statusLabel",
                "Text": "Select a source Text+ clip in the Power Bin / Media Pool, then click Apply.",
                "WordWrap": True,
                "MinimumSize": [500, 48],
            }),

            ui.VGap(8),

            ui.HGroup([
                ui.Button({
                    "ID": "runBtn",
                    "Text": "Apply",
                    "MinimumSize": [160, 36],
                    "Default": True
                }),

                ui.Button({
                    "ID": "closeBtn",
                    "Text": "Close",
                    "MinimumSize": [160, 36]
                }),
            ]),
        ])
    )

    items = window.GetItems()

    for c in COLOR_LIST:
        items["colorCombo"].AddItem(c)

    try:
        items["colorCombo"].CurrentIndex = COLOR_LIST.index(default_color)
    except Exception:
        items["colorCombo"].CurrentIndex = 0

    def read_ui_values():
        try:
            idx = items["colorCombo"].CurrentIndex

            if idx is None or idx < 0 or idx >= len(COLOR_LIST):
                color = DEFAULT_COLOR
            else:
                color = COLOR_LIST[idx]

        except Exception:
            color = DEFAULT_COLOR

        try:
            keep_color = bool(items["keepColorCheck"].Checked)
        except Exception:
            keep_color = True

        preserve_options = {
            "text": bool(items["preserveTextCheck"].Checked),
            "center": bool(items["preserveCenterCheck"].Checked),
            "size": bool(items["preserveSizeCheck"].Checked),
            "angle": bool(items["preserveAngleCheck"].Checked),
            "pivot": bool(items["preservePivotCheck"].Checked),
        }

        try:
            target_track = int(items["trackEdit"].Text)

            if target_track < 0:
                target_track = 0

        except Exception:
            target_track = 0

        return color, keep_color, preserve_options, target_track

    def set_status(text):
        try:
            items["statusLabel"].Text = text
        except Exception:
            print(text)

    def on_run(ev):
        color, keep_color, preserve_options, target_track = read_ui_values()

        save_ui_settings(
            color,
            keep_color,
            preserve_options,
            target_track
        )

        set_status("Processing...")

        try:
            result = run_apply(
                resolve,
                color,
                keep_color,
                preserve_options,
                target_track
            )

            set_status(result["message"])

        except Exception:
            print(traceback.format_exc())
            set_status("An error occurred. Check the console output.")

    def on_close(ev):
        dispatcher.ExitLoop()

    window.On.runBtn.Clicked = on_run
    window.On.closeBtn.Clicked = on_close
    window.On.TextPlusStyleWin.Close = on_close

    window.Show()
    dispatcher.RunLoop()
    window.Hide()


def main():
    print("=== TextPlusStyleCopier ===")

    resolve = dvr.scriptapp("Resolve")

    if not resolve:
        print("Failed to get Resolve instance.")
        return

    show_persistent_ui(resolve)


if __name__ == "__main__":
    main()
