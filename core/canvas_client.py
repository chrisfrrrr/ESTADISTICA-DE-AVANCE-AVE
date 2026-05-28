from __future__ import annotations
import requests
from urllib.parse import urljoin
from typing import Any, Dict, List, Optional, Iterable, Tuple, Union
import time

class CanvasAPIError(Exception):
    pass

class CanvasClient:
    def __init__(self, base_url: str, token: str, timeout: int = 120, max_retries: int = 3):
        self.base_url = base_url.rstrip('/') + '/'
        # Canvas puede tardar bastante cuando se consultan entregas masivas.
        # Usamos timeout separado: conexión corta y lectura amplia.
        self.timeout = (10, timeout)
        self.max_retries = max_retries
        self.session = requests.Session()
        self.session.headers.update({
            'Authorization': f'Bearer {token.strip()}',
            'Accept': 'application/json'
        })

    def _url(self, endpoint: str) -> str:
        endpoint = endpoint.lstrip('/')
        if endpoint.startswith('api/v1/'):
            return urljoin(self.base_url, endpoint)
        return urljoin(self.base_url, 'api/v1/' + endpoint)

    def get(self, endpoint: str, params: Optional[Dict[str, Any]] = None, paginate: bool = True) -> Any:
        url = self._url(endpoint)
        params = dict(params or {})
        params.setdefault('per_page', 100)
        if not paginate:
            r = self._request_with_retry(url, params=params)
            return self._handle(r)
        results: List[Any] = []
        first = True
        while url:
            r = self._request_with_retry(url, params=params if first else None)
            data = self._handle(r)
            if isinstance(data, list):
                results.extend(data)
            else:
                return data
            url = r.links.get('next', {}).get('url')
            first = False
        return results


    def _request_with_retry(self, url: str, params: Optional[Dict[str, Any]] = None) -> requests.Response:
        """Ejecuta GET con reintentos para evitar fallos temporales de Canvas.

        Canvas puede responder lento en cursos grandes, especialmente al extraer
        entregas. Un timeout aislado no debe romper todo el análisis si el
        siguiente intento responde correctamente.
        """
        last_error = None
        for attempt in range(1, self.max_retries + 1):
            try:
                return self.session.get(url, params=params, timeout=self.timeout)
            except requests.exceptions.ReadTimeout as exc:
                last_error = exc
                time.sleep(min(2 * attempt, 6))
            except requests.exceptions.ConnectionError as exc:
                last_error = exc
                time.sleep(min(2 * attempt, 6))
        raise CanvasAPIError(
            'Canvas tardó demasiado en responder. Intente nuevamente o seleccione una sección específica. '
            f'Detalle técnico: {last_error}'
        )


    def post(self, endpoint: str, data: Optional[Union[Dict[str, Any], List[Tuple[str, Any]]]] = None) -> Any:
        """Ejecuta POST contra Canvas.

        Para endpoints como Conversations, Canvas espera parámetros repetidos
        con el nombre exacto `recipients[]`. Por eso este método acepta tanto
        diccionarios como listas de tuplas, que `requests` serializa como:
        recipients[]=1&recipients[]=2.
        """
        url = self._url(endpoint)
        last_error = None
        for attempt in range(1, self.max_retries + 1):
            try:
                r = self.session.post(url, data=data or {}, timeout=self.timeout)
                return self._handle(r)
            except requests.exceptions.ReadTimeout as exc:
                last_error = exc
                time.sleep(min(2 * attempt, 6))
            except requests.exceptions.ConnectionError as exc:
                last_error = exc
                time.sleep(min(2 * attempt, 6))
        raise CanvasAPIError(f'Canvas no respondió al enviar la solicitud POST. Detalle técnico: {last_error}')

    @staticmethod
    def _chunks(values: Iterable[Any], size: int) -> Iterable[List[Any]]:
        chunk = []
        for value in values:
            if value is None:
                continue
            chunk.append(value)
            if len(chunk) >= size:
                yield chunk
                chunk = []
        if chunk:
            yield chunk

    @staticmethod
    def _handle(response: requests.Response) -> Any:
        if response.status_code >= 400:
            try:
                detail = response.json()
            except Exception:
                detail = response.text
            raise CanvasAPIError(f'Canvas respondió {response.status_code}: {detail}')
        if not response.text:
            return None
        return response.json()

    def whoami(self) -> Dict[str, Any]:
        return self.get('users/self', paginate=False)

    def courses(self) -> List[Dict[str, Any]]:
        return self.get('courses', params={'enrollment_state': 'active', 'include[]': ['term']})

    def sections(self, course_id: int | str) -> List[Dict[str, Any]]:
        return self.get(f'courses/{course_id}/sections')

    def enrollments(self, course_id: int | str, section_id: Optional[int | str] = None) -> List[Dict[str, Any]]:
        endpoint = f'sections/{section_id}/enrollments' if section_id else f'courses/{course_id}/enrollments'
        return self.get(endpoint, params={
            'type[]': 'StudentEnrollment',
            'state[]': 'active',
            'include[]': ['user', 'avatar_url']
        })

    def enrollments_for_diagnostic(self, course_id: int | str, section_id: Optional[int | str] = None) -> List[Dict[str, Any]]:
        """Intenta recuperar estudiantes en varios estados para explicar diferencias con Canvas.

        La vista Personas/Módulos de Canvas puede mostrar estudiantes activos, pendientes,
        inactivos o concluidos. El análisis operativo solo usa activos; esta función
        sirve para mostrar cuántos quedaron fuera y por qué, sin afectar el cálculo.
        """
        endpoint = f'sections/{section_id}/enrollments' if section_id else f'courses/{course_id}/enrollments'
        states = ['active', 'invited', 'creation_pending', 'completed', 'inactive']
        found: Dict[int, Dict[str, Any]] = {}
        for state in states:
            try:
                rows = self.get(endpoint, params={
                    'type[]': 'StudentEnrollment',
                    'state[]': state,
                    'include[]': ['user']
                })
                for e in rows or []:
                    u = e.get('user') or {}
                    uid = u.get('id') or e.get('user_id')
                    if uid is not None:
                        e['_diagnostic_state_requested'] = state
                        found[int(uid)] = e
            except Exception:
                continue
        return list(found.values())


    def teachers(self, course_id: int | str) -> List[Dict[str, Any]]:
        return self.get(f'courses/{course_id}/enrollments', params={
            'type[]': ['TeacherEnrollment','TaEnrollment'],
            'state[]': 'active',
            'include[]': ['user']
        })

    def modules(self, course_id: int | str, student_id: Optional[int | str] = None) -> List[Dict[str, Any]]:
        params = {'include[]': ['items', 'content_details']}
        if student_id:
            params['student_id'] = student_id
        modules = self.get(f'courses/{course_id}/modules', params=params)
        # Canvas puede omitir items cuando hay muchos; si pasa, los pedimos por módulo.
        for m in modules or []:
            if 'items' not in m or m.get('items') is None:
                m['items'] = self.module_items(course_id, m.get('id'), student_id=student_id)
        return modules

    def module_items(self, course_id: int | str, module_id: int | str, student_id: Optional[int | str] = None) -> List[Dict[str, Any]]:
        params = {'include[]': ['content_details']}
        if student_id:
            params['student_id'] = student_id
        return self.get(f'courses/{course_id}/modules/{module_id}/items', params=params)

    def find_recipient(self, user_id: int | str, course_id: Optional[int | str] = None) -> List[Dict[str, Any]]:
        """Valida si un usuario es mensajeable por el usuario autenticado.

        Canvas recomienda `/api/v1/search/recipients` para encontrar
        destinatarios válidos. Cuando se pasa `context=course_ID`, limita la
        búsqueda al curso seleccionado.
        """
        params: Dict[str, Any] = {
            'user_id': int(user_id),
            'type': 'user',
        }
        if course_id:
            params['context'] = f'course_{course_id}'
        try:
            data = self.get('search/recipients', params=params, paginate=True)
            return data if isinstance(data, list) else []
        except Exception:
            # Algunas instancias restringen search/recipients. En ese caso no
            # bloqueamos el envío; solamente no validamos previamente.
            return []

    def validate_recipients(self, recipient_ids: List[int], course_id: Optional[int | str] = None) -> Dict[str, List[int]]:
        valid: List[int] = []
        invalid: List[int] = []
        for rid in recipient_ids:
            try:
                result = self.find_recipient(rid, course_id=course_id)
                if result:
                    valid.append(int(rid))
                else:
                    invalid.append(int(rid))
            except Exception:
                invalid.append(int(rid))
        return {'valid': valid, 'invalid': invalid}

    def send_conversation(self, recipient_ids: List[int], subject: str, body: str, course_id: Optional[int | str] = None, group_conversation: bool = False, mode: str = 'async') -> Any:
        """Crea una conversación en Canvas usando el formato correcto.

        Punto importante: Canvas documenta `recipients[]`, no
        `recipients[0]`, `recipients[1]`. En varios tenants, el segundo formato
        puede producir errores 500.
        """
        clean_ids = [int(x) for x in recipient_ids if str(x).isdigit()]
        if not clean_ids:
            raise CanvasAPIError('No hay destinatarios válidos para enviar el mensaje.')

        payload: List[Tuple[str, Any]] = [
            ('subject', subject[:255] if subject else 'Seguimiento académico'),
            ('body', body),
            ('group_conversation', 'true' if group_conversation else 'false'),
            ('force_new', 'true'),
            ('mode', mode if mode in ('sync', 'async') else 'sync'),
        ]
        if course_id:
            payload.append(('context_code', f'course_{course_id}'))
        for rid in clean_ids:
            payload.append(('recipients[]', str(rid)))
        return self.post('conversations', data=payload)

    def send_conversation_safe(
        self,
        recipient_ids: List[int],
        subject: str,
        body: str,
        course_id: Optional[int | str] = None,
        group_conversation: bool = False,
        validate: bool = True,
        chunk_size: int = 20,
    ) -> Dict[str, Any]:
        """Envía mensajes con validación, bloques y fallback individual.

        Devuelve un resumen apto para mostrar en Streamlit sin romper toda la
        ejecución si Canvas rechaza algún destinatario.
        """
        clean_ids = sorted({int(x) for x in recipient_ids if str(x).isdigit()})
        result: Dict[str, Any] = {
            'intentados': len(clean_ids),
            'enviados': 0,
            'omitidos_no_mensajeables': [],
            'errores': [],
            'respuestas_canvas': [],
        }
        if not clean_ids:
            result['errores'].append('No se seleccionaron destinatarios válidos.')
            return result

        ids_to_send = clean_ids
        if validate and course_id:
            checked = self.validate_recipients(clean_ids, course_id=course_id)
            # Si search/recipients está restringido y marca todos inválidos, no
            # bloqueamos el envío; probamos el envío directo con fallback.
            if checked['valid']:
                ids_to_send = checked['valid']
                result['omitidos_no_mensajeables'] = checked['invalid']

        for chunk in self._chunks(ids_to_send, chunk_size):
            try:
                response = self.send_conversation(
                    chunk, subject, body, course_id=course_id,
                    group_conversation=group_conversation,
                    mode='sync' if group_conversation or len(chunk) == 1 else 'async'
                )
                result['enviados'] += len(chunk)
                result['respuestas_canvas'].append(response if response not in (None, []) else {'estado': 'aceptado_en_canvas', 'destinatarios': len(chunk)})
            except Exception as exc:
                # Si falla un bloque, intentamos uno por uno para aislar al
                # destinatario problemático.
                for rid in chunk:
                    try:
                        response = self.send_conversation(
                            [rid], subject, body, course_id=course_id,
                            group_conversation=False, mode='sync'
                        )
                        result['enviados'] += 1
                        result['respuestas_canvas'].append(response if response not in (None, []) else {'estado': 'aceptado_en_canvas', 'destinatario': rid})
                    except Exception as indiv_exc:
                        result['errores'].append({'user_id': rid, 'error': str(indiv_exc)})
        return result

    def assignments(self, course_id: int | str) -> List[Dict[str, Any]]:
        return self.get(f'courses/{course_id}/assignments', params={
            'include[]': ['due_dates', 'all_dates'],
            'order_by': 'due_at'
        })

    def submissions(self, course_id: int | str, student_ids: Optional[List[int]] = None, chunk_size: int = 10) -> List[Dict[str, Any]]:
        """Devuelve entregas de estudiantes.

        En cursos grandes, pedir `student_ids[]=all` puede provocar ReadTimeout
        en Streamlit Cloud. Por eso, cuando tenemos la lista de estudiantes de
        la sección, consultamos en bloques pequeños. Esto hace más solicitudes,
        pero cada una pesa menos y es mucho más estable.
        """
        endpoint = f'courses/{course_id}/students/submissions'

        if student_ids:
            all_results: List[Dict[str, Any]] = []
            clean_ids = [int(x) for x in student_ids if str(x).isdigit()]
            for chunk in self._chunks(clean_ids, chunk_size):
                data = self.get(endpoint, params={
                    'student_ids[]': chunk,
                    'include[]': ['assignment'],
                    'grouped': False
                })
                if isinstance(data, list):
                    all_results.extend(data)
            return all_results

        # Fallback: solo si no hay lista de estudiantes disponible.
        return self.get(endpoint, params={
            'student_ids[]': 'all',
            'include[]': ['assignment'],
            'grouped': False
        })
