-- This annotated SQL file shows how we built up the High Holds query, piece by piece

-- COUNT HOLDS
-- Come up with a number of active title-level holds that aren't suspended ("R*") by bid
SELECT bid, COUNT(bid) as holdcount
FROM transbid_v
WHERE transcode = 'R*'
GROUP BY bid;

-- COUNT ITEMS
-- Pull out a count of "real" items (things that aren't lost, withdrawn, etc.) that are attached to each bib record
--- COUNT requires a GROUP BY, in this case, bid
--- We don't want Lucky Days items to be in mix because they aren't holdable, so we're excluding those two media types (42 and 43)
--- We only want "real" statuses, and we don't want reading station items
SELECT bid, COUNT(bid) as realcount
FROM item_v
WHERE media <> 42 AND media <> 43 
AND (status IN ('C', 'CT', 'H', 'HT', 'I', 'IT', 'IH', 'S', 'ST') OR (status = 'SP' AND owningbranch <> 13))
AND owningbranch <> 2 AND owningbranch <> 11
GROUP BY bid;

-- COUNT COPIES ON ORDER
-- For each order detail, look at the number of copies on order and the number of copies received
SELECT bid, copycount, unitsreceivedcount
FROM orderdetail_v
WHERE status IN ('N', 'O', 'O$', 'R', 'RC', 'RF', 'RG', 'RI', 'RN', 'RP')
AND fund <> '2017ALL';

-- Replace copycount and unitsreceivedcount with an expression to come up with the number of copies actively on order that we're 
-- still waiting for, called pending
SELECT bid, (copycount - unitsreceivedcount) as pending
FROM orderdetail_v
WHERE status IN ('N', 'O', 'O$', 'R', 'RC', 'RF', 'RG', 'RI', 'RN', 'RP')
AND fund <> '2017ALL';

-- Lump together all of the order detail rows together by bib (to account for multiple orders)
SELECT bid, SUM(pending) as onordercount
FROM
    (SELECT bid, (copycount - unitsreceivedcount) as pending
    FROM orderdetail_v
    WHERE status IN ('N', 'O', 'O$', 'R', 'RC', 'RF', 'RG', 'RI', 'RN', 'RP')
    AND fund <> '2017ALL')
GROUP BY bid;

-- PUT TOGETHER BIBS WITH HOLDS, COPIES, ORDERS
--- Without the WHERE condition that leaves out titles without holds, this returns every BID
--- *With* the WHERE condition that leaves out titles without holds, this returns 3,745 bids (as of 6/11/18)
SELECT bibs.bid, holds.holdcount, realitems.realcount, onorder.onordercount
FROM bbibmap_v bibs
LEFT JOIN
    (SELECT bid, COUNT(bid) as holdcount
    FROM transbid_v
    WHERE transcode = 'R*'
    GROUP BY bid) holds
ON bibs.bid = holds.bid
LEFT JOIN
    (SELECT bid, COUNT(bid) as realcount
    FROM item_v
    WHERE media <> 42 AND media <> 43 
    AND (status IN ('C', 'CT', 'H', 'HT', 'I', 'IT', 'IH', 'S', 'ST') OR (status = 'SP' AND owningbranch <> 13))
    AND owningbranch <> 2 AND owningbranch <> 11
    GROUP BY bid) realitems
ON bibs.bid = realitems.bid
LEFT JOIN
    (SELECT bid, SUM(pending) as onordercount
    FROM
        (SELECT bid, (copycount - unitsreceivedcount) as pending
        FROM orderdetail_v
        WHERE status IN ('N', 'O', 'O$', 'R', 'RC', 'RF', 'RG', 'RI', 'RN', 'RP')
        AND fund <> '2017ALL')
    GROUP BY bid) onorder
ON bibs.bid = onorder.bid
WHERE holds.holdcount IS NOT NULL;

-- FILTER OUT BIDs WITH HOLD RATIOS WE DON'T CARE ABOUT (228 bids as of 6/11/18)
-- If we just put in a WHERE clause that says:
--       WHERE holds.holdcount/(realitems.realcount + onorder.onordercount) > 4
-- we won't get what we need, because for bibs that have no items and no orders, it'll try to divide the number of holds by 0
-- and that will return an error. We need to account for that.
-- Making this work involves the CASE function, which takes the form:
--       CASE WHEN the_comparison_for_the_thing_you're_checking is true THEN value ELSE other value END
-- So we want to divide by realcount + onordercount when that doesn't equal zero, and divide by 1 (just for the sake of ease) when it does:
--       CASE WHEN (realitems.realcount + onorder.onordercount) > 0 THEN (realitems.realcount + onorder.onordercount) ELSE 1 END
-- Which becomes:
--       WHERE (holds.holdcount/(CASE WHEN (realitems.realcount + onorder.onordercount) > 0 THEN (realitems.realcount + onorder.onordercount) ELSE 1 END)) > 4
--- (There's actually another several nested sets of CASE statements in the query down below to substitute 0 where the value is null, so the math
--- doesn't go haywire...)
SELECT bibs.bid, holds.holdcount, realitems.realcount, onorder.onordercount
FROM bbibmap_v bibs
LEFT JOIN
    (SELECT bid, COUNT(bid) as holdcount
    FROM transbid_v
    WHERE transcode = 'R*'
    GROUP BY bid) holds
ON bibs.bid = holds.bid
LEFT JOIN
    (SELECT bid, COUNT(bid) as realcount
    FROM item_v
    WHERE media <> 42 AND media <> 43 
    AND (status IN ('C', 'CT', 'H', 'HT', 'I', 'IT', 'IH', 'S', 'ST') OR (status = 'SP' AND owningbranch <> 13))
    AND owningbranch <> 2 AND owningbranch <> 11
    GROUP BY bid) realitems
ON bibs.bid = realitems.bid
LEFT JOIN
    (SELECT bid, SUM(pending) as onordercount
    FROM
        (SELECT bid, (copycount - unitsreceivedcount) as pending
        FROM orderdetail_v
        WHERE status IN ('N', 'O', 'O$', 'R', 'RC', 'RF', 'RG', 'RI', 'RN', 'RP')
        AND fund <> '2017ALL')
    GROUP BY bid) onorder
ON bibs.bid = onorder.bid
WHERE (holds.holdcount/(CASE WHEN ((CASE WHEN realitems.realcount IS NOT NULL THEN realitems.realcount ELSE 0 END) + (CASE WHEN onorder.onordercount IS NOT NULL THEN onorder.onordercount ELSE 0 END)) > 0 THEN ((CASE WHEN realitems.realcount IS NOT NULL THEN realitems.realcount ELSE 0 END) + (CASE WHEN onorder.onordercount IS NOT NULL THEN onorder.onordercount ELSE 0 END)) ELSE 1 END)) > 4
AND holds.holdcount IS NOT NULL;

-- DEAL WITH TITLES THAT HAVE RECEIVED ORDERS, BUT NOT NEW MATCHING ITEMS YET
-- An example as of 6/11/18 is BID 433462 -- we ordered 2 copies and received them against the order, but the new copies aren't 
-- attached to the bib record yet because they haven't been linked. So it looks like we have null for the realcount, and 0 for 
-- the onordercount (because the orders aren't outstanding anymore).
-- We're going to look for bibs that have an item with B itemid attached (which means it's a dummy item created when the order was 
-- placed) and no other items attached.
--- First, count items with B itemids by bid.
SELECT bid, COUNT(item) as receivedcount
FROM item_v
WHERE status IN ('R', 'RC', 'RF')
AND item LIKE 'B%'
GROUP BY bid;

--- Then, total items on the bib.
SELECT bid, COUNT(item) as itemcount
FROM item_v
GROUP BY bid;

--- Combine the two and only return bids with a single item (itemcount = 1) with a B itemid (receivedcount = 1).
SELECT received.bid
FROM 
    (SELECT bid, COUNT(item) as receivedcount
    FROM item_v
    WHERE status IN ('R', 'RC', 'RF')
    AND item LIKE 'B%'
    GROUP BY bid) received
INNER JOIN
    (SELECT bid, COUNT(item) as itemcount
    FROM item_v
    GROUP BY bid) totalitems
ON received.bid = totalitems.bid
WHERE received.receivedcount = 1 AND totalitems.itemcount = 1;

--- Then, we're going to add in the number of pending copies from the order detail information
SELECT received.bid, orders.pendingcount
FROM 
    (SELECT bid, COUNT(item) as receivedcount
    FROM item_v
    WHERE status IN ('R', 'RC', 'RF')
    AND item LIKE 'B%'
    GROUP BY bid) received
INNER JOIN
    (SELECT bid, COUNT(item) as itemcount
    FROM item_v
    GROUP BY bid) totalitems
ON received.bid = totalitems.bid
INNER JOIN
    (SELECT bid, SUM(unitsreceivedcount) as pendingcount
    FROM orderdetail_v
    WHERE destination NOT IN ('BGL-RS', 'WVL-RS')
    AND status IN ('R', 'RC', 'RF')
    GROUP BY bid) orders
ON received.bid = orders.bid
WHERE received.receivedcount = 1 AND totalitems.itemcount = 1;

--- Then, add that *whole* thing in as another LEFT JOIN in the big query
--- (and tack the pending.pendingcount on as another piece of the hold ratio calculation).
SELECT bibs.bid, holds.holdcount, realitems.realcount, onorder.onordercount, pending.pendingcount
FROM bbibmap_v bibs
LEFT JOIN
    (SELECT bid, COUNT(bid) as holdcount
    FROM transbid_v
    WHERE transcode = 'R*'
    GROUP BY bid) holds
ON bibs.bid = holds.bid
LEFT JOIN
    (SELECT bid, COUNT(bid) as realcount
    FROM item_v
    WHERE media <> 42 AND media <> 43 
    AND (status IN ('C', 'CT', 'H', 'HT', 'I', 'IT', 'IH', 'S', 'ST') OR (status = 'SP' AND owningbranch <> 13))
    AND owningbranch <> 2 AND owningbranch <> 11
    GROUP BY bid) realitems
ON bibs.bid = realitems.bid
LEFT JOIN
    (SELECT bid, SUM(pending) as onordercount
    FROM
        (SELECT bid, (copycount - unitsreceivedcount) as pending
        FROM orderdetail_v
        WHERE status IN ('N', 'O', 'O$', 'R', 'RC', 'RF', 'RG', 'RI', 'RN', 'RP')
        AND fund <> '2017ALL')
    GROUP BY bid) onorder
ON bibs.bid = onorder.bid
LEFT JOIN
    (SELECT received.bid, orders.pendingcount
    FROM 
        (SELECT bid, COUNT(item) as receivedcount
        FROM item_v
        WHERE status IN ('R', 'RC', 'RF')
        AND item LIKE 'B%'
        GROUP BY bid) received
    INNER JOIN
        (SELECT bid, COUNT(item) as itemcount
        FROM item_v
        GROUP BY bid) totalitems
    ON received.bid = totalitems.bid
    INNER JOIN
        (SELECT bid, SUM(unitsreceivedcount) as pendingcount
        FROM orderdetail_v
        WHERE destination NOT IN ('BGL-RS', 'WVL-RS')
        AND status IN ('R', 'RC', 'RF')
        GROUP BY bid) orders
    ON received.bid = orders.bid
    WHERE received.receivedcount = 1 AND totalitems.itemcount = 1) pending
ON bibs.bid = pending.bid
WHERE (holds.holdcount/(CASE WHEN ((CASE WHEN realitems.realcount IS NOT NULL THEN realitems.realcount ELSE 0 END) + (CASE WHEN onorder.onordercount IS NOT NULL THEN onorder.onordercount ELSE 0 END) + (CASE WHEN pending.pendingcount IS NOT NULL THEN pending.pendingcount ELSE 0 END)) > 0 THEN ((CASE WHEN realitems.realcount IS NOT NULL THEN realitems.realcount ELSE 0 END) + (CASE WHEN onorder.onordercount IS NOT NULL THEN onorder.onordercount ELSE 0 END) + (CASE WHEN pending.pendingcount IS NOT NULL THEN pending.pendingcount ELSE 0 END)) ELSE 1 END)) > 4
AND holds.holdcount IS NOT NULL;


-- THE FULL QUERY
-- We're going to add on a couple of other fields and tables, so that we have title, author, and format for when we view the list.
-- We're also going add one more CASE statement right at the beginning, to return either then number of copies on order, or the number
-- that are pending.
SELECT bibs.bid, bibs.author, bibs.title, bibs.isbn, format.formattext, holds.holdcount, realitems.realcount, 
    (CASE WHEN (onorder.onordercount = 0 OR onorder.onordercount IS NULL) THEN pending.pendingcount ELSE onorder.onordercount END) as ordercount
FROM bbibmap_v bibs
LEFT JOIN
    (SELECT bid, COUNT(bid) as holdcount
    FROM transbid_v
    WHERE transcode = 'R*'
    GROUP BY bid) holds
ON bibs.bid = holds.bid
LEFT JOIN
    (SELECT bid, COUNT(bid) as realcount
    FROM item_v
    WHERE media <> 42 AND media <> 43 
    AND (status IN ('C', 'CT', 'H', 'HT', 'I', 'IT', 'IH', 'S', 'ST') OR (status = 'SP' AND owningbranch <> 13))
    AND owningbranch <> 2 AND owningbranch <> 11
    GROUP BY bid) realitems
ON bibs.bid = realitems.bid
LEFT JOIN
    (SELECT bid, SUM(pending) as onordercount
    FROM
        (SELECT bid, (copycount - unitsreceivedcount) as pending
        FROM orderdetail_v
        WHERE status IN ('N', 'O', 'O$', 'R', 'RC', 'RF', 'RG', 'RI', 'RN', 'RP')
        AND fund <> '2017ALL')
    GROUP BY bid) onorder
ON bibs.bid = onorder.bid
LEFT JOIN
    (SELECT received.bid, orders.pendingcount
    FROM 
        (SELECT bid, COUNT(item) as receivedcount
        FROM item_v
        WHERE status IN ('R', 'RC', 'RF')
        AND item LIKE 'B%'
        GROUP BY bid) received
    INNER JOIN
        (SELECT bid, COUNT(item) as itemcount
        FROM item_v
        GROUP BY bid) totalitems
    ON received.bid = totalitems.bid
    INNER JOIN
        (SELECT bid, SUM(unitsreceivedcount) as pendingcount
        FROM orderdetail_v
        WHERE destination NOT IN ('BGL-RS', 'WVL-RS')
        AND status IN ('R', 'RC', 'RF')        
        GROUP BY bid) orders
    ON received.bid = orders.bid
    WHERE received.receivedcount = 1 AND totalitems.itemcount = 1) pending
ON bibs.bid = pending.bid
LEFT JOIN
formatterm_v format
ON bibs.format = format.formattermid
WHERE (holds.holdcount/(CASE WHEN ((CASE WHEN realitems.realcount IS NOT NULL THEN realitems.realcount ELSE 0 END) + (CASE WHEN onorder.onordercount IS NOT NULL THEN onorder.onordercount ELSE 0 END) + (CASE WHEN pending.pendingcount IS NOT NULL THEN pending.pendingcount ELSE 0 END)) > 0 THEN ((CASE WHEN realitems.realcount IS NOT NULL THEN realitems.realcount ELSE 0 END) + (CASE WHEN onorder.onordercount IS NOT NULL THEN onorder.onordercount ELSE 0 END) + (CASE WHEN pending.pendingcount IS NOT NULL THEN pending.pendingcount ELSE 0 END)) ELSE 1 END)) > 4
AND holds.holdcount IS NOT NULL;

-- This full query gets pulled into Google by script, which goes through and recalculates hold ratio based on format.
