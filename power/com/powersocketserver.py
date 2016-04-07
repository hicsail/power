#!/usr/bin/python
# -*- coding: utf-8 -*-

import json
import logging as log

from power.com import Signal

from gevent.lock import BoundedSemaphore
from geventwebsocket import WebSocketServer, WebSocketApplication, Resource
from power import config
from pydispatch import dispatcher


def build_message(status, message=None, state=None):
    """
    Build message string to send to connected clients.

    :param status: HTTP status code of the message, int
    :param message: Optional free-form message to display to the client, can be string or dictionary.
        Either this or state should be set.
    :param state: Optional state to pass to the client. This can be the full state by passing in config.data() or a
        subset that you define yourself. Should be a dictionary. Either this or message should be set.
    :return: JSON string
    """
    # The server needs to send either a message or state to inform other clients what's going on.
    # The client sending a command will know that its action was successful and can update its UI, but the other
    # clients would only know that something was successful, not what.
    if message is None and state is None:
        raise ValueError('Either  message or state should be set')
    msg = {'status': status, 'message': message, 'state': state}
    return json.dumps(msg)


class PowerSocketServer(WebSocketApplication):
    """
    A socket server to communicate UI updates to connected clients, and receive commands to control PW.

    The following data format should be followed:
    For the client to the server: { “action”: “foo”, “value”: 1 } - Value is optional
    For the server to the client: {“status”: 200, “message”: “Some message”, “state”: {}} - Message and state are
    optional, but not both at the same time.
    """
    sem = BoundedSemaphore(1)

    def __init__(self, ws):
        self.paused = config.get('Paused', False)
        # Register update UI signal
        dispatcher.connect(self.handle_ui_update, signal=Signal.UPDATE_UI_SIGNAL, sender=dispatcher.Any)
        super().__init__(ws)

    def on_open(self):
        """ Client connected handler, send new client application state """
        log.info('Client connected')
        # Send full config whenever a client connects
        self.ws.send(build_message(200, state=config.data()))

    def on_message(self, message, **kwargs):
        """
        Message received handler

        :param message: Message received from client, of type str
        :param kwargs:
        """
        if message is None:
            return
        # Check for invalid JSON
        try:
            message = json.loads(message)
        except ValueError:
            # Let the client know about it (only the one who sent it)
            self.ws.send(build_message(400, 'Invalid JSON'))
        else:
            # Check for pause
            self.check_pause(message)
            # Send message to poweralgorithm.py
            dispatcher.send(signal=Signal.PW_COMMAND_SIGNAL, message=message)

    def broadcast(self, message):
        """
        Broadcast message to all connected clients

        :param message: Message to broadcast, of type dictionary
        """
        for client in self.ws.handler.server.clients.values():
            client.ws.send(message)

    def on_close(self, reason):
        """
        Client disconnected handler

        :param reason: Reason for closing socket
        """
        print('Connection closed: %s' % reason)

    def handle_ui_update(self, sender, message):
        """
        Event handler for UI updates

        :param sender: Signal sender, not used
        :param message: Message to broadcast to connected clients, must be constructed using build_message
        """
        if message:
            self.broadcast(message)

    def check_pause(self, message):
        """
        Intercept message before passing on to PW class and check if pause/resume command is present

        :param message: Client message
        """
        try:
            if message['command'] == 'pause':
                self.pause()
            if message['command'] == 'resume':
                self.resume()
        except KeyError:
            # Command is not set, nothing to do
            pass

    def pause(self):
        """
        Pause execution of PW
        """
        if not self.paused:
            # Acquire lock, power.add_task won't run anymore
            PowerSocketServer.sem.acquire()
            log.info('Pause PW')
            self.paused = True
            config.put('paused', True)
            self.ws.send(build_message(200, state={'paused': 1}))

    def resume(self):
        """
        Resume execution of PW
        """
        if self.paused:
            # Release lock, power.add_task will continue
            PowerSocketServer.sem.release()
            log.info('Resume PW')
            self.paused = False
            config.put('paused', False)
            self.ws.send(build_message(200, state={'paused': 0}))


def init():
    WebSocketServer(
            ('127.0.0.1', config.get('Port', 7000)),
            Resource([
                ('^/socket', PowerSocketServer)
            ]),
            debug=config.get('DebugSocketServer', 0)
    ).serve_forever()


if __name__ == "__main__":
    init()
