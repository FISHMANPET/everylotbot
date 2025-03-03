# -*- coding: utf-8 -*-
# This file is part of everylotbot
# Copyright 2016 Neil Freeman
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import argparse
import logging

# import twitter_bot_utils as tbu
# from . import __version__ as version
from everylot import EveryLot
from bsky_client import init_client
from atproto import Client


def main():
    parser = argparse.ArgumentParser(description="every lot twitter bot")
    parser.add_argument(
        "screen_name",
        metavar="SCREEN_NAME",
        type=str,
        help="Twitter screen name (without @)",
        nargs="?",
        const=1,
        default="every-usps-box",
    )
    parser.add_argument(
        "database",
        metavar="DATABASE",
        type=str,
        help="path to SQLite lots database",
        nargs="?",
        const=1,
        default="lots.db",
    )
    parser.add_argument(
        "--id",
        type=str,
        default=None,
        help="tweet the entry in the lots table with this id",
    )
    parser.add_argument(
        "-s",
        "--search-format",
        type=str,
        default=None,
        metavar="STRING",
        help="Python format string use for searching Google",
    )
    parser.add_argument(
        "-p",
        "--print-format",
        type=str,
        default=None,
        metavar="STRING",
        help="Python format string use for poster to Twitter",
    )
    parser.add_argument(
        "-n", "--dry-run", action="store_true", help="Don't actually do anything"
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Run talkatively")
    parser.add_argument("-q", "--quiet", action="store_true", help="Run quietly")
    # tbu.args.add_default_args(parser, version=version, include=('config', 'dry-run', 'verbose', 'quiet'))

    args = parser.parse_args()
    # api = tbu.api.API(args)
    client = init_client()

    logging.basicConfig(filename="%s.log" % args.screen_name)
    logger = logging.getLogger(args.screen_name)
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    elif args.quiet:
        logger.setLevel(logging.ERROR)
    else:
        logger.setLevel(logging.INFO)
    logger.debug("everylot starting with %s, %s", args.screen_name, args.database)
    find_lots(args, client, logger)


def find_lots(args, client: Client, logger):
    els = [
        EveryLot(
            args.database,
            logger=logger,
            print_format=args.print_format,
            search_format=args.search_format,
            id_=None,
        )
        for _ in range(4)
    ]
    if args.id:
        els[0] = EveryLot(
            args.database,
            logger=logger,
            print_format=args.print_format,
            search_format=args.search_format,
            id_=args.id,
        )

    if not els[0].lot:
        logger.error("No lots found")
        return

    for el in els:
        logger.debug(
            "%s addresss: %s zip: %s",
            el.lot["id"],
            el.lot.get("address"),
            el.lot.get("zip5"),
        )
        logger.debug("db location %s,%s", el.lot["lat"], el.lot["lon"])

    # for now, if it doesn't have imagery, we're gonna just mark it in the db
    # as "posted = 1" skip the process and give me
    # a chance to troubleshoot!
    with open("streetview.txt") as f:
        sv_key = f.read()

    # metadata = el.get_streetview_metadata(sv_key)
    # if metadata is False:
    #     logger.error("No imagery going on here :(")
    #     el.mark_as_no_imagery()
    #     if args.id:
    #         return
    #     else:
    #         find_lot(args, api, logger)
    #         return

    # Get the streetview image and upload it
    # ("sv.jpg" is a dummy value, since filename is a required parameter).
    images = [el.get_streetview_image(sv_key) for el in els]
    # media = api.media_upload('sv.jpg', file=image)

    # compose an update with all the good parameters
    # including the media string.
    updates = [el.compose() for el in els]
    for update in updates:
        logger.info(update["status"])

    if not args.dry_run:
        logger.debug("posting")
        update = "\n".join([u["status"] for u in updates])
        if len(updates) > 300:
            logger.error("Update too long, trying again")
            find_lots(args, client, logger)

        image_alts = [f"Street view image of {el.lot["name"].title()}" for el in els]

        status = client.send_images(
            text=update,
            images=images,
            image_alts=image_alts,
        )
        for el in els:
            try:
                el.mark_as_posted(status.uri)
            except AttributeError:
                el.mark_as_posted("1")


if __name__ == "__main__":
    main()
