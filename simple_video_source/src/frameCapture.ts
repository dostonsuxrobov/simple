/** Wait for a pending seek to decode, without changing the user's playhead. */
export async function waitForPresentedFrame(video: HTMLVideoElement, intendedTime: number) {
  await new Promise<void>((resolve, reject) => {
    let settled = false
    let frameId = 0
    const cleanup = () => {
      clearTimeout(timer)
      cancelAnimationFrame(frameId)
      video.removeEventListener('seeked', check)
      video.removeEventListener('loadeddata', check)
      video.removeEventListener('error', failed)
      video.removeEventListener('emptied', failed)
    }
    const finish = (error?: Error) => {
      if (settled) return
      settled = true
      cleanup()
      if (error) reject(error)
      else resolve()
    }
    const failed = () => finish(new Error('The video changed or could not decode this frame. Try again.'))
    const check = () => {
      if (settled || video.seeking || video.readyState < HTMLMediaElement.HAVE_CURRENT_DATA) return
      cancelAnimationFrame(frameId)
      frameId = requestAnimationFrame(() => {
        frameId = requestAnimationFrame(() => {
          if (video.seeking || video.readyState < HTMLMediaElement.HAVE_CURRENT_DATA) return
          if (Math.abs(video.currentTime - intendedTime) > 0.002) {
            finish(new Error('The playhead moved while capturing. Try again at the selected frame.'))
          } else finish()
        })
      })
    }
    const timer = setTimeout(() => finish(new Error('This frame is still loading. Wait for playback, then try again.')), 5_000)
    video.addEventListener('seeked', check)
    video.addEventListener('loadeddata', check)
    video.addEventListener('error', failed)
    video.addEventListener('emptied', failed)
    check()
  })
}
