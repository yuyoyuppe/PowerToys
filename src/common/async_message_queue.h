#pragma once
#include <queue>
#include <thread>
#include <mutex>
#include <condition_variable>
#include <string>

class AsyncMessageQueue {
private:
  std::mutex queue_mutex;
  std::queue<std::wstring> message_queue;
  std::condition_variable message_ready;

  //Disable copy
  AsyncMessageQueue(const AsyncMessageQueue&);
  AsyncMessageQueue& operator=(const AsyncMessageQueue&);

public:
  AsyncMessageQueue() {
  }
  void queue_message(std::wstring message) {
    this->queue_mutex.lock();
    this->message_queue.push(message);
    this->queue_mutex.unlock();
    this->message_ready.notify_one();
  }
  std::wstring pop_message() {
    std::unique_lock<std::mutex> lock(this->queue_mutex);
    while (message_queue.empty()) {
      this->message_ready.wait(lock);
    }
    std::wstring message = this->message_queue.front();
    this->message_queue.pop();
    return message;
  }
};
