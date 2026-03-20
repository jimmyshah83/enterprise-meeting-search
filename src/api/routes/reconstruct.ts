/**
 * Reconstruction API routes.
 *
 * POST /api/v1/reconstruct/:meetingId — trigger meeting reconstruction
 */

import { Router, Request, Response, NextFunction } from 'express';
import { v4 as uuidv4 } from 'uuid';
import { searchStore, reconstructionStore } from '../store';
import { reconstructionRateLimiter } from '../middleware/rateLimiter';
import { NotFoundError } from '../middleware/errorHandler';
import { MeetingReconstructionEngine } from '../../reconstruction/MeetingReconstructionEngine';

const router = Router();
const engine = new MeetingReconstructionEngine();

/**
 * @openapi
 * /api/v1/reconstruct/{meetingId}:
 *   post:
 *     summary: Trigger meeting reconstruction from search results
 *     description: >
 *       Reconstructs a structured meeting summary (title, date/time, attendees,
 *       topics, decisions, action items) from any previously collected search
 *       results that reference the given meeting ID.
 *     tags: [Reconstruction]
 *     parameters:
 *       - in: path
 *         name: meetingId
 *         required: true
 *         schema:
 *           type: string
 *         description: The meeting ID to reconstruct
 *     responses:
 *       202:
 *         description: Reconstruction triggered successfully
 *       404:
 *         description: Meeting ID not found in any completed search results
 *       429:
 *         description: Rate limit exceeded
 *       500:
 *         description: Internal server error
 */
router.post('/:meetingId', reconstructionRateLimiter, async (req: Request, res: Response, next: NextFunction) => {
  try {
    const { meetingId } = req.params;

    // Find the meeting result in any completed searches
    const meetingResult = searchStore.findMeetingResult(meetingId);
    if (!meetingResult) {
      return next(new NotFoundError(
        `Meeting '${meetingId}' was not found in any completed search results. ` +
        'Run a search first to populate meeting data.',
      ));
    }

    const reconstructionId = uuidv4();
    reconstructionStore.set({
      reconstructionId,
      meetingId,
      status: 'running',
      createdAt: new Date(),
    });

    // Perform reconstruction asynchronously
    Promise.resolve()
      .then(() => engine.reconstruct(meetingResult))
      .then((result) => {
        reconstructionStore.update(reconstructionId, {
          status: 'completed',
          result,
          completedAt: new Date(),
        });
      })
      .catch((err: unknown) => {
        reconstructionStore.update(reconstructionId, {
          status: 'failed',
          error: err instanceof Error ? err.message : 'Unknown error',
          completedAt: new Date(),
        });
      });

    // Wait briefly for synchronous reconstruction to complete (it's fast)
    await new Promise((resolve) => setTimeout(resolve, 50));
    const record = reconstructionStore.get(reconstructionId);

    if (record?.status === 'completed') {
      return res.status(200).json({
        reconstructionId,
        meetingId,
        status: 'completed',
        reconstruction: record.result,
        completedAt: record.completedAt,
      });
    }

    return res.status(202).json({
      reconstructionId,
      meetingId,
      status: 'running',
      message: 'Reconstruction triggered. Check back shortly for results.',
    });
  } catch (err) {
    next(err);
  }
});

export default router;
