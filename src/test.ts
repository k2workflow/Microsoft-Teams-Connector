import test from 'ava';
import '@k2oss/k2-broker-core/test-framework';
import './index';


//
// WARNING: any tests that use this mock() must be run serially
// (using test.serial()), because they modify global  !
function mock(name: string, value: any) 
{
    global[name] = value;
}

// helper method to pass the data which is dynamic to each test into the mocked XHR class object.
// uses mock() so must only be used in test.serial()'y run tests
function mockXHR(data){
    let xhr: {[key:string]: any} = null;
    class XHR {
        public onreadystatechange: () => void;
        public readyState: number;
        public status: number;
        public responseText: string;
        private recorder: {[key:string]: any};

        constructor() {
            xhr = this.recorder = {};
            this.recorder.headers = {};
        }

        open(method: string, url: string) {
            this.recorder.opened = {method, url};
        }

        setRequestHeader(key: string, value: string) {
            this.recorder.headers[key] = value;
        }

        send() {
            queueMicrotask(() =>
            {
                this.readyState = 4;
                this.status = 200;
                // this.responseText = JSON.stringify({
                //     "id": 4321,
                //     "requestStatusUrl": undefined,
                //     "isSuccessful": true
                // });
                this.responseText = JSON.stringify(data);
                this.onreadystatechange();
                delete this.responseText;
            });
        }
    }

    mock('XMLHttpRequest', XHR);
}

test('onexecute fails for invalid object', t => {
    t.throws(function () {
        let obj = 'invalidObject';
        onexecute(obj, '', {}, {});
    });
});

//
// // example of how to catch a throw exception
// const promise = () => Promise.reject(new Error('TEST'));
// test('rejects', async t => {
//     const error = await t.throwsAsync(promise);
//     t.is(error.message, 'TEST');
// });

test.serial('onexecuteTeamArchive succeeds', async t => {
    //
    // note: inputs/outputs found index.ts ondescribe "postSchema()" method

    let validObject = Team;
    let method = TeamArchive;

    let data = {
        [TeamId]: 1234
    };
    mockXHR(data);

    let result: any = null;
    function pr(r: any) {
        result = r;
    }

    mock('postResult', pr);

    await onexecute(validObject, method, {}, data);

    // t.deepEqual(xhr, {
    //     opened: {
    //         method: 'GET',
    //         url: 'https://jsonplaceholder.typicode.com/todos/123'
    //     },
    //     headers: {
    //         'test': 'test value'
    //     }
    // });

    t.deepEqual(result, {
        [TeamId]: 1234,
        [TeamRequestStatusUrl]: undefined,
        [TeamIsSuccessful]: true
    });

    t.pass();
});

test.serial('onexecuteTeamUnarchive succeeds', async t => {
    //
    // note: inputs/outputs found index.ts ondescribe "postSchema()" method

    let teamId = 4321;
    let validObject = Team;
    let method = TeamUnarchive;

    let data = {
        [TeamId]: teamId
    };
    mockXHR(data);

    let result: any = null;
    function pr(r: any) { result = r;  }
    mock('postResult', pr);

    await onexecute(validObject, method, {}, data);

    t.deepEqual(result, {
        [TeamId]: teamId,
        [TeamRequestStatusUrl]: undefined,
        [TeamIsSuccessful]: true
    });

    t.pass();
});

test.serial('onexecuteTeamClone succeeds', async t => {
    //
    // note: inputs/outputs found index.ts ondescribe "postSchema()" method

    let validObject = Team;
    let method = TeamClone;

    let teamId = 999;
    let data = {
        [TeamId]: teamId,
        [TeamDisplayName]: "SomeName",
        [TeamDescription]: "SomeDescription",
        [TeamMailNickname]: "SomeMailNickname"
    };
    mockXHR(data);

    let result: any = null;
    function pr(r: any) { result = r;  }
    mock('postResult', pr);

    await onexecute(validObject, method, {}, data);

    t.deepEqual(result,{
        [TeamId]: teamId,
        [TeamRequestStatusUrl]: undefined,
        [TeamIsSuccessful]: true
    });

    t.pass();
});

test.serial('onexecuteTeamUpdate succeeds', async t => {
    //
    // note: inputs/outputs found index.ts ondescribe "postSchema()" method

    let validObject = Team;
    let method = TeamUpdate;

    let teamId = 555;
    let data = {
        [TeamId]: teamId,
        [TeamMsAllowCreateUpdateChannels]: true,
        [TeamMsAllowDeleteChannels]: true,
        [TeamMsAllowAddRemoveApps]: true,
        [TeamMsAllowCreateUpdateRemoveTabs]: true,
        [TeamMsAllowCreateUpdateRemoveConnectors]: true,
        [TeamGsAllowCreateUpdateChannels]: true,
        [TeamGsAllowDeleteChannels]: true,
        [TeamMsgAllowUserEditMessages]: true,
        [TeamMsgAllowUserDeleteMessages]: true,
        [TeamMsgAllowTeamMentions]: true,
        [TeamMsgAllowChannelMentions]: true,
        [TeamFsAllowGiphy]: true,
        [TeamFsAllowStickersAndMemes]: true,
        [TeamFsAllowCustomMemes]: true
    };
    mockXHR(data);

    let result: any = null;
    function pr(r: any) { result = r;  }
    mock('postResult', pr);

    await onexecute(validObject, method, {}, data);

    t.deepEqual(result,{
        [TeamIsSuccessful]: true
    });

    t.pass();
});

test.serial('onexecuteChannelUpdate succeeds', async t => {
    //
    // note: inputs/outputs found index.ts ondescribe "postSchema()" method

    let validObject = Channel;
    let method = ChannelUpdate;

    let channelId = 777;
    let teamId = 888;
    let data = {
        [ChannelId]: channelId,
        [ChannelTeamId]: teamId,
        [ChannelDisplayName]: "SomeName",
        [ChannelDescription]: "SomeDescription"
    };
    mockXHR(data);

    let result: any = null;
    function pr(r: any) { result = r;  }
    mock('postResult', pr);

    await onexecute(validObject, method, {}, data);

    t.deepEqual(result,{
        [ChannelIsSuccessful]: true
    });

    t.pass();
});

test.serial('onexecuteChannelDelete succeeds', async t => {
    //
    // note: inputs/outputs found index.ts ondescribe "postSchema()" method

    let validObject = Channel;
    let method = ChannelDelete;

    let channelId = 345;
    let teamId = 100;
    let data = {
        [ChannelId]: channelId,
        [ChannelTeamId]: teamId
    };
    mockXHR(data);

    let result: any = null;
    function pr(r: any) { result = r;  }
    mock('postResult', pr);

    await onexecute(validObject, method, {}, data);

    t.deepEqual(result,{
        [ChannelIsSuccessful]: true
    });

    t.pass();
});

test.serial('onexecuteSendMessage (of a Channel) succeeds', async t => {
    //
    // note: inputs/outputs found index.ts ondescribe "postSchema()" method

    let validObject = Channel;
    let method = ChannelSendMessage;

    let channelId = 345;
    let teamId = 100;
    let data = {
        [ChannelId]: channelId,
        [ChannelTeamId]: teamId,
        [ChannelMessageSubject]: "subject",
        [ChannelMessageBody]: "body",
        [ChannelMessageIsImportant]: true
    };
    mockXHR(data);

    let result: any = null;
    function pr(r: any) { result = r;  }
    mock('postResult', pr);

    await onexecute(validObject, method, {}, data);

    t.deepEqual(result,{
        [ChannelIsSuccessful]: true
    });

    t.pass();
});