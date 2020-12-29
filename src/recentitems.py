# Keypirinha launcher (keypirinha.com)

import keypirinha as kp
import keypirinha_util as kpu
from keypirinha_wintypes import FOLDERID

import os.path
import time


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

    scan_recent_directory = DEFAULT_SCAN_RECENT_DIRECTORY

    def __init__(self):
        super().__init__()

    def on_start(self):
        self._read_config()

    def on_catalog(self):
        catalog = []

        if self.scan_recent_directory:
            start = time.perf_counter()

            self._add_from_recent_directory(catalog)

            self._log_catalog_duration(start, "link", len(catalog))

        self.set_catalog(catalog)

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

        config_changed = new_scan_recent_directory != self.scan_recent_directory
        self.scan_recent_directory = new_scan_recent_directory

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

    def _log_catalog_duration(self, start, what, count):
        elapsed = time.perf_counter() - start
        self.info("Cataloged {} {}{} in {:.1f} seconds".format(count, what, "s"[count == 1:], elapsed))
