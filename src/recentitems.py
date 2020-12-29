# Keypirinha launcher (keypirinha.com)

import keypirinha as kp
import keypirinha_util as kpu
from keypirinha_wintypes import FOLDERID

import os.path
import time
import winreg as reg


def _enum_registry_keys(key, sub_key):
    try:
        with reg.OpenKey(key, sub_key) as key:
            count = reg.QueryInfoKey(key)[0]
            for i in range(count):
                yield os.path.join(sub_key, reg.EnumKey(key, i))
    except OSError:
        return []


class RecentItems(kp.Plugin):
    """
    One-line description of your plugin.

    This block is a longer and more detailed description of your plugin that may
    span on several lines, albeit not being required by the application.

    You may have several plugins defined in this module. It can be useful to
    logically separate the features of your package. All your plugin classes
    will be instantiated by Keypirinha as long as they are derived directly or
    indirectly from :py:class:`keypirinha.Plugin` (aliased ``kp.Plugin`` here).

    In case you want to have a base class for your plugins, you must prefix its
    name with an underscore (``_``) to indicate Keypirinha it is not meant to be
    instantiated directly.

    In rare cases, you may need an even more powerful way of telling Keypirinha
    what classes to instantiate: the ``__keypirinha_plugins__`` global variable
    may be declared in this module. It can be either an iterable of class
    objects derived from :py:class:`keypirinha.Plugin`; or, even more dynamic,
    it can be a callable that returns an iterable of class objects. Check out
    the ``StressTest`` example from the SDK for an example.

    Up to 100 plugins are supported per module.

    More detailed documentation at: http://keypirinha.com/api/plugin.html
    """

    CONFIG_SECTION_MAIN = "main"

    DEFAULT_SCAN_RECENT_DIRECTORY = True
    DEFAULT_SCAN_RECENT_ITEMS = True

    scan_recent_directory = DEFAULT_SCAN_RECENT_DIRECTORY
    scan_recent_items = DEFAULT_SCAN_RECENT_ITEMS

    def __init__(self):
        super().__init__()

    def on_start(self):
        self._read_config()

    def on_catalog(self):
        if self.scan_recent_directory:
            start = time.perf_counter()
            catalog = []

            self._add_from_recent_directory(catalog)
            self.merge_catalog(catalog)

            self._log_catalog_duration(start, "link", len(catalog))

        if self.scan_recent_items:
            start = time.perf_counter()
            catalog = []

            self._add_from_recent_items(catalog)
            self.merge_catalog(catalog)

            self._log_catalog_duration(start, "item", len(catalog))

    def on_execute(self, item, action):
        kpu.execute_default_action(self, item, action)

    def on_events(self, flags):
        if flags & kp.Events.PACKCONFIG:
            if self._read_config():
                self.on_catalog()

    def _read_config(self):
        new_scan_recent_directory = self.load_settings().get_bool(
            "scan_recent_directory", self.CONFIG_SECTION_MAIN, self.DEFAULT_SCAN_RECENT_DIRECTORY
        )

        new_scan_recent_items = self.load_settings().get_bool(
            "scan_recent_items", self.CONFIG_SECTION_MAIN, self.DEFAULT_SCAN_RECENT_ITEMS
        )

        config_changed = (
            new_scan_recent_directory != self.scan_recent_directory or
            new_scan_recent_items != self.scan_recent_items
        )

        self.scan_recent_directory = new_scan_recent_directory
        self.scan_recent_items = new_scan_recent_items

        return config_changed

    def _add_from_recent_directory(self, catalog):
        recent_directory = os.path.join(
            kpu.shell_known_folder_path(FOLDERID.RoamingAppData.value),
            "Microsoft", "Windows", "Recent"
        )

        try:
            recent_links = kpu.scan_directory(
                recent_directory, "*.lnk",
                kpu.ScanFlags.FILES,
                max_level=0
            )
        except OSError as e:
            self.dbg(str(e))
            return

        for link_name in recent_links:
            link_path = os.path.join(recent_directory, link_name)
            item = self.create_item(
                category=kp.ItemCategory.FILE,
                label=os.path.splitext(link_name)[0],
                short_desc="Recent link: " + link_path,
                target=link_path,
                args_hint=kp.ItemArgsHint.ACCEPTED,
                hit_hint=kp.ItemHitHint.KEEPALL
            )
            catalog.append(item)

    def _add_from_recent_items(self, catalog):
        try:
            root_key = reg.HKEY_CURRENT_USER
            recent_apps_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Search\RecentApps"
            for app_key in _enum_registry_keys(root_key, recent_apps_key):
                for item_key in _enum_registry_keys(root_key, os.path.join(app_key, "RecentItems")):
                    with reg.OpenKey(root_key, item_key) as item:
                        path = reg.QueryValueEx(item, "Path")[0]
                        item = self.create_item(
                            category=kp.ItemCategory.FILE,
                            label=os.path.basename(path),
                            short_desc="Recent item: " + path,
                            target=path,
                            args_hint=kp.ItemArgsHint.ACCEPTED,
                            hit_hint=kp.ItemHitHint.KEEPALL
                        )
                        catalog.append(item)
        except OSError as e:
            self.dbg(str(e))
            return

    def _log_catalog_duration(self, start, what, count):
        elapsed = time.perf_counter() - start
        self.info("Cataloged {} {}{} in {:.1f} seconds".format(count, what, "s"[count == 1:], elapsed))
