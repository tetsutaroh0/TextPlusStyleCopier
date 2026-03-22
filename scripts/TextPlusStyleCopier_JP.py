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


def get_settings_file_path():
    try:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "ApplyTextPlusStyle_settings.json")
    except Exception:
        return "ApplyTextPlusStyle_settings.json"


def load_ui_settings():
    settings_path = get_settings_file_path()
    default_settings = {
        "last_color": DEFAULT_COLOR,
        "keep_color": True
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
        return data
    except Exception:
        return default_settings


def save_ui_settings(color, keep_color):
    settings_path = get_settings_file_path()
    data = {
        "last_color": color if color in COLOR_LIST else DEFAULT_COLOR,
        "keep_color": bool(keep_color)
    }
    try:
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("設定保存に失敗:", e)


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


def get_all_timeline_video_items(timeline):
    result = []
    try:
        track_count = timeline.GetTrackCount("video")
    except Exception:
        return result

    for track_index in range(1, track_count + 1):
        try:
            items = timeline.GetItemListInTrack("video", track_index) or []
        except Exception:
            items = []
        result.extend(items)

    return result


def append_clip_to_timeline_and_detect_new_item(media_pool, timeline, media_pool_item):
    before_items = get_all_timeline_video_items(timeline)
    before_ids = set(id(x) for x in before_items)

    try:
        appended = media_pool.AppendToTimeline([media_pool_item])
    except Exception as e:
        print("AppendToTimeline に失敗:", e)
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


def copy_style_keep_text(src_tool, dst_comp, dst_tool):
    try:
        src_settings = src_tool.SaveSettings()
    except Exception as e:
        print("SaveSettings に失敗:", e)
        return False

    try:
        dst_text_before = dst_tool.GetInput("StyledText")
    except Exception:
        dst_text_before = None

    try:
        dst_comp.Lock()
    except Exception:
        pass

    ok = True

    try:
        dst_tool.LoadSettings(src_settings)
    except Exception as e:
        print("LoadSettings 中にエラー:", e)
        ok = False

    if dst_text_before is not None:
        try:
            dst_tool.SetInput("StyledText", dst_text_before)
        except Exception as e:
            print("StyledText 復元中にエラー:", e)
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


def find_colored_textplus_items(timeline, target_color, exclude_item=None):
    result = []
    for item in get_all_timeline_video_items(timeline):
        if exclude_item and item == exclude_item:
            continue

        color = get_clip_color_safe(item)
        if color != target_color:
            continue

        if is_textplus_timeline_item(item):
            result.append(item)

    return result


def run_apply(resolve, target_color, keep_color):
    result = {
        "ok": False,
        "message": "",
        "copied_count": 0,
        "failed_count": 0,
    }

    pm = resolve.GetProjectManager()
    project = pm.GetCurrentProject() if pm else None
    if not project:
        result["message"] = "プロジェクトが開かれていません。"
        return result

    media_pool = project.GetMediaPool()
    timeline = project.GetCurrentTimeline()
    if not timeline:
        result["message"] = "タイムラインが開かれていません。"
        return result

    original_timecode = get_current_timecode_safe(timeline)

    src_media_pool_item = get_selected_mediapool_clip(media_pool)
    if not src_media_pool_item:
        result["message"] = "Power Bin / Media Pool で参照元の Text+ を1つ選択してください。"
        return result

    ref_timeline_item = append_clip_to_timeline_and_detect_new_item(media_pool, timeline, src_media_pool_item)
    if not ref_timeline_item:
        result["message"] = "参照元クリップをタイムラインへ追加できませんでした。"
        return result

    try:
        set_current_timecode_safe(timeline, project, original_timecode)

        src_comp, src_tool = get_textplus_from_timeline_item(ref_timeline_item)
        if not src_tool:
            result["message"] = "参照元クリップ内に TextPlus ツールが見つかりません。"
            return result

        target_items = find_colored_textplus_items(timeline, target_color, exclude_item=ref_timeline_item)
        if not target_items:
            result["message"] = f"{target_color} の Text+ クリップが見つかりません。"
            return result

        copied_count = 0
        failed_count = 0

        for item in target_items:
            dst_comp, dst_tool = get_textplus_from_timeline_item(item)
            if not dst_tool:
                failed_count += 1
                continue

            ok = copy_style_keep_text(src_tool, dst_comp, dst_tool)
            if ok:
                copied_count += 1
                if not keep_color:
                    clear_clip_color_safe(item)
            else:
                failed_count += 1

        result["ok"] = True
        result["copied_count"] = copied_count
        result["failed_count"] = failed_count
        result["message"] = f"完了: 成功 {copied_count} / 失敗 {failed_count}"
        return result

    finally:
        delete_timeline_item_safe(timeline, ref_timeline_item)
        set_current_timecode_safe(timeline, project, original_timecode)


def show_persistent_ui(resolve):
    if bmd is None:
        print("UIManager 用の BlackmagicFusion が読み込めませんでした。")
        return

    saved_settings = load_ui_settings()
    default_color = saved_settings.get("last_color", DEFAULT_COLOR)
    default_keep_color = bool(saved_settings.get("keep_color", True))

    try:
        fusion = resolve.Fusion()
        ui = fusion.UIManager
        dispatcher = bmd.UIDispatcher(ui)
    except Exception as e:
        print("UI 初期化に失敗:", e)
        return

    window = dispatcher.AddWindow(
        {
            "ID": "TextPlusStyleWin",
            "WindowTitle": "Apply Text+ Style",
            "Geometry": [100, 100, 560, 270],
        },
        ui.VGroup(
            [
                ui.Label({
                    "ID": "colorLabel",
                    "Text": "対象クリップカラー",
                    "Alignment": {"AlignLeft": True, "AlignVCenter": True},
                }),
                ui.ComboBox({
                    "ID": "colorCombo",
                    "MinimumSize": [500, 34],
                    "MaximumSize": [16777215, 34],
                }),
                ui.VGap(4),
                ui.CheckBox({
                    "ID": "keepColorCheck",
                    "Text": "処理後もクリップカラーを残す",
                    "Checked": default_keep_color,
                }),
                ui.VGap(8),
                ui.Label({
                    "ID": "statusLabel",
                    "Text": "参照元 Text+ を Power Bin / Media Pool で選択してから実行してください。",
                    "WordWrap": True,
                    "Alignment": {"AlignLeft": True, "AlignTop": True},
                    "MinimumSize": [500, 40],
                }),
                ui.VGap(8),
                ui.HGroup(
                    [
                        ui.Button({
                            "ID": "runBtn",
                            "Text": "実行",
                            "MinimumSize": [160, 36],
                            "Default": True
                        }),
                        ui.Button({
                            "ID": "closeBtn",
                            "Text": "閉じる",
                            "MinimumSize": [160, 36]
                        }),
                    ]
                ),
            ]
        ),
    )

    items = window.GetItems()

    for c in COLOR_LIST:
        items["colorCombo"].AddItem(c)

    try:
        items["colorCombo"].CurrentIndex = COLOR_LIST.index(default_color)
    except Exception:
        try:
            items["colorCombo"].CurrentIndex = COLOR_LIST.index(DEFAULT_COLOR)
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

        return color, keep_color

    def set_status(text):
        try:
            items["statusLabel"].Text = text
        except Exception:
            print(text)

    def on_run(ev):
        color, keep_color = read_ui_values()
        save_ui_settings(color, keep_color)
        set_status("処理中...")

        try:
            result = run_apply(resolve, color, keep_color)
            set_status(result["message"])
        except Exception as e:
            print(traceback.format_exc())
            set_status("エラーが発生しました。コンソールを確認してください。")

    def on_close(ev):
        dispatcher.ExitLoop()

    window.On.runBtn.Clicked = on_run
    window.On.closeBtn.Clicked = on_close
    window.On.TextPlusStyleWin.Close = on_close

    window.Show()
    dispatcher.RunLoop()
    window.Hide()


def main():
    print("=== Apply Text+ Style with Persistent UI ===")

    resolve = dvr.scriptapp("Resolve")
    if not resolve:
        print("Resolve インスタンスを取得できませんでした。")
        return

    show_persistent_ui(resolve)


if __name__ == "__main__":
    main()