# 論文存檔用
citation排名指令SELECT (SELECT `title` FROM `paper_sort_out` WHERE `id` = `p_id`)AS 'title', (SELECT `author_key` FROM `paper_sort_out` WHERE `id` = `p_id`)AS 'keyword',(SELECT `citation` FROM `paper_sort_out` WHERE `id` = `p_id`)AS 'citation' FROM `paper_keyword` WHERE `k_id` = 24 OR `k_id` = 3 OR `k_id` = 630 OR `k_id` = 1635 ORDER BY `citation`  DESC
