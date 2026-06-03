# CBDB User MDB — Issues Report

_A respectful summary of issues uncovered during regression testing._

Dear maintainer,

Below is a summary of the issues we uncovered while building an automated regression-test suite for the CBDB User MDB. We hope this report is useful as you continue your wonderful stewardship of this dataset, and we sincerely thank you for the immense work that has gone into building it.

The issues are ordered by severity (P0 highest). Each entry includes a concise description, step-by-step user reproduction, screenshots where the issue is visible in the Access UI, and a suggested fix. None of these are urgent; they are documented so they can be addressed at the maintainer's convenience.

## Table of Contents

- [P2 — Silent display](#p2--silent-display)
  - [Issue #1 — LookAtEntry: c_entry_desc backfill is NULL for all rows when entry_code = 36 (jinshi general)](#issue-1--lookatentry-c_entry_desc-backfill-is-null-for-all-rows-when-entry_code--36-jinshi-general)
  - [Issue #2 — LookAtGroupData: CmdRun does not backfill c_name from BIOG_MAIN](#issue-2--lookatgroupdata-cmdrun-does-not-backfill-c_name-from-biog_main)
- [Severity legend](#severity-legend)
- [Appendix — c_index_year / c_index_addr_id drift vs the cbdb-online-main-server snapshot (differences need per-row classification before being filed as bugs)](#appendix--c_index_year--c_index_addr_id-drift-vs-the-cbdb-online-main-server-snapshot-differences-need-per-row-classification-before-being-filed-as-bugs)
- [Closing note](#closing-note)

## Severity legend

- P0 — Silent data corruption: data is wrong or missing without an error popup.
- P1 — Visible runtime crash: a popup appears, the operation aborts.
- P2 — Silent display: form fields render blank when they should show data.
- P3 — Missing UI: a feature exists in code but no button invokes it.
- P4 — Setup: one-time hurdle on each new install.
- P5 — Dormant / latent / not currently reproducible: kept as historical record; we re-checked on the current dump and could not trigger the symptom.

## P2 — Silent display

### Issue #1 — LookAtEntry: c_entry_desc backfill is NULL for all rows when entry_code = 36 (jinshi general)

**Affected sub:** `Form_LookAtEntry.CmdQuery_Click`

**Severity:** P2 — Silent display issue: 92,545 rows affected.  The user can see the blank c_entry_desc column in the result grid, but Access shows no error — making it easy to overlook.  Exports (GIS, Neo4j, KML) that reference this column will also carry the blank.

#### Description

When the user runs a LookAtEntry query filtered to entry code 36 (examination: jinshi general), the result table ZZ_SCRATCH_ENTRY is populated with 92,545 rows but the c_entry_desc column is NULL for every row.  The expected value is 'examination: jinshi (general)'.

The CmdQuery_Click handler successfully inserts rows from ENTRY_DATA joined to ENTRY_CODES, but the c_entry_desc backfill step does not write the description for this specific entry code.  All other columns appear to be filled normally.  The missing description means the on-screen result grid shows a blank entry-type column for every record, which is misleading — the user sees results but cannot identify what type of examination each record represents.

Detected by: test_vba_full_matrix[top_entry_code_36_unfiltered] — assertion 'c_entry_desc backfill wrong' with 92,545 affected rows.

#### Steps to reproduce

1. 1
2. .
3.  
4. O
5. p
6. e
7. n
8.  
9. C
10. B
11. D
12. B
13. _
14. B
15. J
16. _
17. U
18. s
19. e
20. r
21. .
22. m
23. d
24. b
25.  
26. i
27. n
28.  
29. M
30. i
31. c
32. r
33. o
34. s
35. o
36. f
37. t
38.  
39. A
40. c
41. c
42. e
43. s
44. s
45. .
46. 

47. 2
48. .
49.  
50. F
51. r
52. o
53. m
54.  
55. t
56. h
57. e
58.  
59. N
60. a
61. v
62. i
63. g
64. a
65. t
66. i
67. o
68. n
69.  
70. P
71. a
72. n
73. e
74. ,
75.  
76. o
77. p
78. e
79. n
80.  
81. t
82. h
83. e
84.  
85. f
86. o
87. r
88. m
89.  
90. *
91. *
92. L
93. o
94. o
95. k
96. A
97. t
98. E
99. n
100. t
101. r
102. y
103. *
104. *
105. .
106. 

107. 3
108. .
109.  
110. I
111. n
112.  
113. t
114. h
115. e
116.  
117. E
118. n
119. t
120. r
121. y
122.  
123. C
124. o
125. d
126. e
127.  
128. p
129. i
130. c
131. k
132. e
133. r
134. ,
135.  
136. s
137. e
138. l
139. e
140. c
141. t
142.  
143. e
144. n
145. t
146. r
147. y
148.  
149. c
150. o
151. d
152. e
153.  
154. *
155. *
156. 3
157. 6
158. *
159. *
160.  
161. (
162. l
163. a
164. b
165. e
166. l
167. :
168.  
169. '
170. e
171. x
172. a
173. m
174. i
175. n
176. a
177. t
178. i
179. o
180. n
181. :
182.  
183. j
184. i
185. n
186. s
187. h
188. i
189.  
190. (
191. g
192. e
193. n
194. e
195. r
196. a
197. l
198. )
199. '
200. )
201. .
202. 

203. 4
204. .
205.  
206. L
207. e
208. a
209. v
210. e
211.  
212. d
213. y
214. n
215. a
216. s
217. t
218. y
219. ,
220.  
221. a
222. d
223. d
224. r
225. e
226. s
227. s
228. ,
229.  
230. a
231. n
232. d
233.  
234. y
235. e
236. a
237. r
238.  
239. f
240. i
241. l
242. t
243. e
244. r
245. s
246.  
247. b
248. l
249. a
250. n
251. k
252. .
253. 

254. 5
255. .
256.  
257. C
258. l
259. i
260. c
261. k
262.  
263. *
264. *
265. R
266. u
267. n
268.  
269. Q
270. u
271. e
272. r
273. y
274. *
275. *
276.  
277. (
278. C
279. m
280. d
281. Q
282. u
283. e
284. r
285. y
286.  
287. b
288. u
289. t
290. t
291. o
292. n
293. )
294. .
295. 

296. 6
297. .
298.  
299. W
300. h
301. e
302. n
303.  
304. t
305. h
306. e
307.  
308. q
309. u
310. e
311. r
312. y
313.  
314. c
315. o
316. m
317. p
318. l
319. e
320. t
321. e
322. s
323. ,
324.  
325. i
326. n
327. s
328. p
329. e
330. c
331. t
332.  
333. t
334. h
335. e
336.  
337. r
338. e
339. s
340. u
341. l
342. t
343.  
344. g
345. r
346. i
347. d
348. :
349.  
350. t
351. h
352. e
353.  
354. e
355. n
356. t
357. r
358. y
359. -
360. t
361. y
362. p
363. e
364.  
365. d
366. e
367. s
368. c
369. r
370. i
371. p
372. t
373. i
374. o
375. n
376.  
377. c
378. o
379. l
380. u
381. m
382. n
383.  
384. (
385. c
386. _
387. e
388. n
389. t
390. r
391. y
392. _
393. d
394. e
395. s
396. c
397. )
398.  
399. i
400. s
401.  
402. b
403. l
404. a
405. n
406. k
407.  
408. f
409. o
410. r
411.  
412. e
413. v
414. e
415. r
416. y
417.  
418. r
419. o
420. w
421. .
422. 

423. 7
424. .
425.  
426. S
427. Q
428. L
429.  
430. v
431. e
432. r
433. i
434. f
435. i
436. c
437. a
438. t
439. i
440. o
441. n
442. :
443.  
444. `
445. S
446. E
447. L
448. E
449. C
450. T
451.  
452. T
453. O
454. P
455.  
456. 5
457.  
458. c
459. _
460. e
461. n
462. t
463. r
464. y
465. _
466. c
467. o
468. d
469. e
470. ,
471.  
472. c
473. _
474. e
475. n
476. t
477. r
478. y
479. _
480. d
481. e
482. s
483. c
484.  
485. F
486. R
487. O
488. M
489.  
490. Z
491. Z
492. _
493. S
494. C
495. R
496. A
497. T
498. C
499. H
500. _
501. E
502. N
503. T
504. R
505. Y
506. `
507.  
508. r
509. e
510. t
511. u
512. r
513. n
514. s
515.  
516. (
517. 3
518. 6
519. ,
520.  
521. N
522. U
523. L
524. L
525. )
526.  
527. f
528. o
529. r
530.  
531. a
532. l
533. l
534.  
535. r
536. o
537. w
538. s
539. .

#### Suggested fix

Locate the backfill step in Form_LookAtEntry.CmdQuery_Click that sets c_entry_desc for ZZ_SCRATCH_ENTRY rows.  Verify that the JOIN to ENTRY_CODES on c_entry_code = 36 is not inadvertently filtered out or that the UPDATE / backfill SQL matches the column name exactly.  After the fix, `SELECT c_entry_desc FROM ZZ_SCRATCH_ENTRY WHERE c_entry_code = 36 LIMIT 1` should return 'examination: jinshi (general)'.

### Issue #2 — LookAtGroupData: CmdRun does not backfill c_name from BIOG_MAIN

**Affected sub:** `Form_LookAtGroupData.CmdRun_Click`

**Severity:** P2 — Silent display issue: CmdRun completes without any error message, but the c_name column in the result is blank.  The user has no indication that the backfill failed.

#### Description

When the user seeds a person ID into LookAtGroupData and clicks CmdRun, the handler is expected to run an UPDATE query that joins ZZ_SCRATCH_IMPORT_PEOPLE to BIOG_MAIN and fills in c_name (and c_dynasty) for each seeded row.  In this build the UPDATE does not execute successfully: after CmdRun completes, c_name remains NULL in ZZ_SCRATCH_IMPORT_PEOPLE.

The result is that the group-data import display shows empty name cells.  The user has no indication that the backfill failed — CmdRun does not surface an error.

Detected by: test_hard_form_query_small_fixture[groupdata_person_1_small] — assertion 'CmdRun didn't backfill c_name for c_person_id=1', c_name is None after CmdRun completes.

#### Steps to reproduce

1. 1
2. .
3.  
4. O
5. p
6. e
7. n
8.  
9. C
10. B
11. D
12. B
13. _
14. B
15. J
16. _
17. U
18. s
19. e
20. r
21. .
22. m
23. d
24. b
25.  
26. i
27. n
28.  
29. M
30. i
31. c
32. r
33. o
34. s
35. o
36. f
37. t
38.  
39. A
40. c
41. c
42. e
43. s
44. s
45. .
46. 

47. 2
48. .
49.  
50. F
51. r
52. o
53. m
54.  
55. t
56. h
57. e
58.  
59. N
60. a
61. v
62. i
63. g
64. a
65. t
66. i
67. o
68. n
69.  
70. P
71. a
72. n
73. e
74. ,
75.  
76. o
77. p
78. e
79. n
80.  
81. t
82. h
83. e
84.  
85. f
86. o
87. r
88. m
89.  
90. *
91. *
92. L
93. o
94. o
95. k
96. A
97. t
98. G
99. r
100. o
101. u
102. p
103. D
104. a
105. t
106. a
107. *
108. *
109. .
110. 

111. 3
112. .
113.  
114. I
115. n
116.  
117. t
118. h
119. e
120.  
121. i
122. m
123. p
124. o
125. r
126. t
127.  
128. p
129. e
130. r
131. s
132. o
133. n
134.  
135. l
136. i
137. s
138. t
139. ,
140.  
141. e
142. n
143. t
144. e
145. r
146.  
147. a
148.  
149. v
150. a
151. l
152. i
153. d
154.  
155. p
156. e
157. r
158. s
159. o
160. n
161.  
162. I
163. D
164.  
165. (
166. e
167. .
168. g
169. .
170.  
171. *
172. *
173. 1
174. *
175. *
176. )
177. .
178. 

179. 4
180. .
181.  
182. C
183. l
184. i
185. c
186. k
187.  
188. *
189. *
190. R
191. u
192. n
193. *
194. *
195.  
196. (
197. C
198. m
199. d
200. R
201. u
202. n
203.  
204. b
205. u
206. t
207. t
208. o
209. n
210. )
211. .
212. 

213. 5
214. .
215.  
216. W
217. h
218. e
219. n
220.  
221. C
222. m
223. d
224. R
225. u
226. n
227.  
228. c
229. o
230. m
231. p
232. l
233. e
234. t
235. e
236. s
237. ,
238.  
239. i
240. n
241. s
242. p
243. e
244. c
245. t
246.  
247. t
248. h
249. e
250.  
251. r
252. e
253. s
254. u
255. l
256. t
257. :
258.  
259. t
260. h
261. e
262.  
263. N
264. a
265. m
266. e
267.  
268. c
269. o
270. l
271. u
272. m
273. n
274.  
275. i
276. s
277.  
278. b
279. l
280. a
281. n
282. k
283. .
284. 

285. 6
286. .
287.  
288. S
289. Q
290. L
291.  
292. v
293. e
294. r
295. i
296. f
297. i
298. c
299. a
300. t
301. i
302. o
303. n
304. :
305.  
306. `
307. S
308. E
309. L
310. E
311. C
312. T
313.  
314. c
315. _
316. p
317. e
318. r
319. s
320. o
321. n
322. _
323. i
324. d
325. ,
326.  
327. c
328. _
329. n
330. a
331. m
332. e
333.  
334. F
335. R
336. O
337. M
338.  
339. Z
340. Z
341. _
342. S
343. C
344. R
345. A
346. T
347. C
348. H
349. _
350. I
351. M
352. P
353. O
354. R
355. T
356. _
357. P
358. E
359. O
360. P
361. L
362. E
363. `
364.  
365. r
366. e
367. t
368. u
369. r
370. n
371. s
372.  
373. (
374. 1
375. ,
376.  
377. N
378. U
379. L
380. L
381. )
382.  
383. —
384.  
385. c
386. _
387. n
388. a
389. m
390. e
391.  
392. w
393. a
394. s
395.  
396. n
397. o
398. t
399.  
400. b
401. a
402. c
403. k
404. f
405. i
406. l
407. l
408. e
409. d
410.  
411. f
412. r
413. o
414. m
415.  
416. B
417. I
418. O
419. G
420. _
421. M
422. A
423. I
424. N
425. .

#### Suggested fix

Locate the UPDATE statement in Form_LookAtGroupData.CmdRun_Click that joins ZZ_SCRATCH_IMPORT_PEOPLE to BIOG_MAIN and sets c_name. Check that the JOIN condition matches the correct key column and that the UPDATE target column name is spelled correctly.  After the fix, running CmdRun with any valid person ID should populate c_name in ZZ_SCRATCH_IMPORT_PEOPLE.

## Closing note

Thank you for taking the time to read this report. None of the items above is urgent; we hope having them all in one place makes it easy to address them at your own pace.

If any of the descriptions or suggested fixes are unclear, we would be glad to discuss further. The corresponding regression tests in this repository will automatically flip from PASS to FAIL the moment any regression marker stops reproducing in the source dump — that is a signal to investigate, not an automatic confirmation that the bug is fixed (the marker could fail because of an upstream fix, a fixture / driver change on our side, or a misclassification we made earlier).
