import { jest } from '@jest/globals'
import {
	buildProfilePictureQueryContent,
	buildTextStatusUpdateContent,
	executeTextStatusUpdate
} from '../../Socket/chats'
import type { ILogger } from '../../Utils/logger'
import type { BinaryNode } from '../../WABinary'

const mockLogger: ILogger = {
	level: 'silent',
	trace: jest.fn(),
	debug: jest.fn(),
	info: jest.fn(),
	warn: jest.fn(),
	error: jest.fn(),
	child: () => mockLogger
}

describe('buildProfilePictureQueryContent', () => {
	it('builds a profile picture query without tctoken content', () => {
		expect(buildProfilePictureQueryContent('preview')).toEqual([
			{
				tag: 'picture',
				attrs: { type: 'preview', query: 'url' }
			}
		])
	})

	it('nests tctoken under the picture node', () => {
		const tcToken: BinaryNode = {
			tag: 'tctoken',
			attrs: { t: '1770000000' },
			content: Buffer.from([4, 1, 33])
		}

		expect(buildProfilePictureQueryContent('image', [tcToken])).toEqual([
			{
				tag: 'picture',
				attrs: { type: 'image', query: 'url' },
				content: [tcToken]
			}
		])
	})
})

describe('buildTextStatusUpdateContent', () => {
	it('formats plain string input with default 0 ephemeral duration', () => {
		const result = buildTextStatusUpdateContent('Available')
		expect(result).toEqual({
			mexInput: {
				text: 'Available',
				ephemeral_duration_sec: 0
			},
			legacyStatusText: 'Available'
		})
	})

	it('formats object input with emoji and ephemeral duration', () => {
		const result = buildTextStatusUpdateContent({
			text: 'Working remotely',
			emoji: '🏠',
			ephemeralDuration: 86400
		})
		expect(result).toEqual({
			mexInput: {
				text: 'Working remotely',
				emoji: { content: '🏠' },
				ephemeral_duration_sec: 86400
			},
			legacyStatusText: '🏠 Working remotely'
		})
	})

	it('handles empty string or whitespace clearing text status', () => {
		const resultEmpty = buildTextStatusUpdateContent('')
		expect(resultEmpty).toEqual({
			mexInput: {
				text: null,
				ephemeral_duration_sec: 0
			},
			legacyStatusText: ''
		})

		const resultWhitespace = buildTextStatusUpdateContent('   ')
		expect(resultWhitespace).toEqual({
			mexInput: {
				text: null,
				ephemeral_duration_sec: 0
			},
			legacyStatusText: ''
		})
	})

	it('formats emoji-only status', () => {
		const result = buildTextStatusUpdateContent({
			emoji: '🚀'
		})
		expect(result).toEqual({
			mexInput: {
				text: null,
				emoji: { content: '🚀' },
				ephemeral_duration_sec: 0
			},
			legacyStatusText: '🚀'
		})
	})
})

describe('executeTextStatusUpdate', () => {
	it('updates both MEX text status and legacy status when a string is passed', async () => {
		const queriesSent: BinaryNode[] = []
		const mockQuery = jest.fn(async (node: BinaryNode) => {
			queriesSent.push(node)
			if (node.attrs.xmlns === 'w:mex') {
				return {
					tag: 'iq',
					attrs: { type: 'result' },
					content: [
						{
							tag: 'result',
							attrs: {},
							content: Buffer.from(
								JSON.stringify({
									data: {
										xwa2_update_text_status: { result: true }
									}
								}),
								'utf-8'
							)
						}
					]
				} as BinaryNode
			}

			return { tag: 'iq', attrs: { type: 'result' } } as BinaryNode
		})

		await executeTextStatusUpdate('Listening to: Redbone — Childish Gambino', mockQuery, () => 'tag_123', mockLogger)

		expect(mockQuery).toHaveBeenCalledTimes(2)

		// 1. Verify MEX query
		const mexQuery = queriesSent[0]!
		expect(mexQuery.attrs.xmlns).toBe('w:mex')
		const queryChild = (mexQuery.content as BinaryNode[])[0]!
		expect(queryChild.attrs.query_id).toBe('9152604461510864')
		const variables = JSON.parse(queryChild.content!.toString())
		expect(variables).toEqual({
			variables: {
				input: {
					text: 'Listening to: Redbone — Childish Gambino',
					ephemeral_duration_sec: 0
				}
			}
		})

		// 2. Verify legacy status query
		const legacyQuery = queriesSent[1]!
		expect(legacyQuery.attrs.xmlns).toBe('status')
		const statusChild = (legacyQuery.content as BinaryNode[])[0]!
		expect(statusChild.content!.toString()).toBe('Listening to: Redbone — Childish Gambino')
	})

	it('supports TextStatusInput with emoji and ephemeralDuration', async () => {
		const queriesSent: BinaryNode[] = []
		const mockQuery = jest.fn(async (node: BinaryNode) => {
			queriesSent.push(node)
			if (node.attrs.xmlns === 'w:mex') {
				return {
					tag: 'iq',
					attrs: { type: 'result' },
					content: [
						{
							tag: 'result',
							attrs: {},
							content: Buffer.from(
								JSON.stringify({
									data: {
										xwa2_update_text_status: { result: true }
									}
								}),
								'utf-8'
							)
						}
					]
				} as BinaryNode
			}

			return { tag: 'iq', attrs: { type: 'result' } } as BinaryNode
		})

		await executeTextStatusUpdate(
			{
				text: 'Busy working',
				emoji: '💻',
				ephemeralDuration: 7200
			},
			mockQuery,
			() => 'tag_456',
			mockLogger
		)

		expect(mockQuery).toHaveBeenCalledTimes(2)

		// 1. Verify MEX query payload
		const mexQuery = queriesSent[0]!
		const queryChild = (mexQuery.content as BinaryNode[])[0]!
		const variables = JSON.parse(queryChild.content!.toString())
		expect(variables).toEqual({
			variables: {
				input: {
					text: 'Busy working',
					emoji: { content: '💻' },
					ephemeral_duration_sec: 7200
				}
			}
		})

		// 2. Verify legacy status payload joins emoji and text
		const legacyQuery = queriesSent[1]!
		const statusChild = (legacyQuery.content as BinaryNode[])[0]!
		expect(statusChild.content!.toString()).toBe('💻 Busy working')
	})

	it('handles clearing status when empty string is passed', async () => {
		const queriesSent: BinaryNode[] = []
		const mockQuery = jest.fn(async (node: BinaryNode) => {
			queriesSent.push(node)
			if (node.attrs.xmlns === 'w:mex') {
				return {
					tag: 'iq',
					attrs: { type: 'result' },
					content: [
						{
							tag: 'result',
							attrs: {},
							content: Buffer.from(
								JSON.stringify({
									data: {
										xwa2_update_text_status: { result: true }
									}
								}),
								'utf-8'
							)
						}
					]
				} as BinaryNode
			}

			return { tag: 'iq', attrs: { type: 'result' } } as BinaryNode
		})

		await executeTextStatusUpdate('', mockQuery, () => 'tag_789', mockLogger)

		const mexQuery = queriesSent[0]!
		const queryChild = (mexQuery.content as BinaryNode[])[0]!
		const variables = JSON.parse(queryChild.content!.toString())
		expect(variables).toEqual({
			variables: {
				input: {
					text: null,
					ephemeral_duration_sec: 0
				}
			}
		})

		const legacyQuery = queriesSent[1]!
		const statusChild = (legacyQuery.content as BinaryNode[])[0]!
		expect(statusChild.content!.toString()).toBe('')
	})

	it('does not throw when secondary legacy update fails', async () => {
		const mockQuery = jest.fn(async (node: BinaryNode) => {
			if (node.attrs.xmlns === 'w:mex') {
				return {
					tag: 'iq',
					attrs: { type: 'result' },
					content: [
						{
							tag: 'result',
							attrs: {},
							content: Buffer.from(
								JSON.stringify({
									data: {
										xwa2_update_text_status: { result: true }
									}
								}),
								'utf-8'
							)
						}
					]
				} as BinaryNode
			}

			throw new Error('legacy status failed')
		})

		await expect(executeTextStatusUpdate('Hello', mockQuery, () => 'tag_err', mockLogger)).resolves.not.toThrow()
	})

	it('throws error when primary MEX query fails', async () => {
		const mockQuery = jest.fn(async (node: BinaryNode) => {
			if (node.attrs.xmlns === 'w:mex') {
				return {
					tag: 'iq',
					attrs: { type: 'error' },
					content: [
						{
							tag: 'error',
							attrs: { code: '500', text: 'internal error' }
						}
					]
				} as BinaryNode
			}

			return { tag: 'iq', attrs: { type: 'result' } } as BinaryNode
		})

		await expect(executeTextStatusUpdate('Hello', mockQuery, () => 'tag_err', mockLogger)).rejects.toThrow()
	})
})
