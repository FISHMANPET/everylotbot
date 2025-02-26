from atproto import Client, Session, SessionEvent
import logging

logging.basicConfig('everylot_usps.log')
logger = logging.getLogger('bsky_client')
logger.setLevel(logging.INFO)


def get_session() -> str:
  with open('session.txt') as f:
    return f.read()


def save_session(session_string: str) -> None:
  with open('session.txt', 'w') as f:
    f.write(session_string)


def on_session_change(event: SessionEvent, session: Session) -> None:
  logger.debug('Session changed:', event, repr(session))
  if event in (SessionEvent.CREATE, SessionEvent.REFRESH):
    logger.info('Saving changed session')
    save_session(session.export())


def init_client() -> Client:
  client = Client()
  client.on_session_change(on_session_change)
  session_string = get_session()
  logger.debug('Reusing session')
  client.login(session_string=session_string)

  return client
