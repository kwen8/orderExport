<template>
  <Row type="flex" justify="center" align="middle" class="login__layout">
    <Col>
      <Card style="width:350px">
        <p slot="title">
          <Icon type="ios-film-outline"></Icon>
          登录
        </p>
        <Form ref="formCustom" :model="loginForm" :rules="ruleCustom" :label-width="80">
          <FormItem label="用户名" prop="user">
            <Input type="text" v-model="loginForm.user" />
          </FormItem>
          <FormItem label="密码" prop="password">
            <Input type="password" v-model="loginForm.password" />
          </FormItem>
          <FormItem>
            <Button type="primary" @click="handleSubmit('loginForm')">登录</Button>
            <Button type="ghost" @click="handleReset('loginForm')" style="margin-left: 8px">重置</Button>
          </FormItem>
        </Form>
      </Card>
    </Col>
  </Row>
</template>

<script>
export default {
  data () {
    return {
      loginForm: {
        user: '',
        password: ''
      }
    }
  }
}
</script>

<style lang="scss" scope>
  .login__layout {
    height: 100%;
  }
</style>

